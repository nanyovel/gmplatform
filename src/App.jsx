import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signOut,
} from "firebase/auth";
import {
  collection,
  doc,
  getDoc,
  getDocs,
  onSnapshot,
  setDoc,
  writeBatch,
} from "firebase/firestore";
import {
  ArrowDown,
  ArrowUp,
  ArrowUpDown,
  ChevronDown,
  Download,
  LogOut,
  Pin,
  PinOff,
  Plus,
  Share2,
  Upload,
} from "lucide-react";
import { ADMINS, auth, db } from "./firebase";

// Normaliza nombres: sin tildes, mayúsculas, asteriscos ni comas. "short" quita (paréntesis) y S.R.L.
const norm = (s) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[*,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
const short = (s) =>
  norm(
    String(s)
      .replace(/\(.*?\)/g, "")
      .replace(/\bs\.?r\.?l\b\.?/gi, ""),
  );
const keys = (s) => [norm(s), short(s)];
const clean = (s) => s.replace(/\*$/, "");
const blank = () => ["", "", false];
const blank2 = () => ["", ""];
const YES = /^(x|si|sí|true|1|verdadero|prioritario)$/i;
const inp = "rounded border border-slate-300 bg-white px-3 py-1.5 text-sm";
const MAX_PINS = 3; // máximo de almacenes que se pueden fijar a la vez
const MAX_STORES = 5; // máximo de tiendas en "Exhibido en hasta"
// Comparador para ordenar columnas: números por valor, textos alfabéticos (sin tildes ni mayúsculas, "2" antes que "10")
const coll = new Intl.Collator("es", { numeric: true, sensitivity: "base" });
const by = (x, y) => (typeof x === "number" ? x - y : coll.compare(x, y));
const fdate = new Intl.DateTimeFormat("es-DO", {
  dateStyle: "medium",
  timeStyle: "short",
});
// Prioritarios guardados con el formato antiguo [código, sucursal 1, 2, 3, todas] pasan a una fila [código, sucursal, todas] por sucursal
const toRows = (rows) =>
  rows.flatMap((r) => {
    if (r.length < 5) return [r];
    const [c, a, b, d, all] = r;
    if (all) return [[c, "", true]];
    const names = [a, b, d].map((x) => x.trim()).filter(Boolean);
    return names.length ? names.map((x) => [c, x, false]) : [[c, "", false]];
  });
// config/main guardaba un solo almacén fijado (texto); ahora guarda una lista de hasta 3
const asList = (p) => (Array.isArray(p) ? p : p ? [p] : []);

// Descarga un .xlsx con exactamente estas columnas y filas
const download = (name, head, body) => {
  const ws = XLSX.utils.aoa_to_sheet([head, ...body]);
  ws["!cols"] = head.map((h, c) => ({
    wch: Math.min(
      50,
      Math.max(
        String(h).length,
        ...body.slice(0, 300).map((r) => String(r[c] ?? "").length),
      ) + 2,
    ),
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, name);
  const d = new Date(),
    p = (n) => String(n).padStart(2, "0");
  XLSX.writeFile(
    wb,
    `${name} ${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.xlsx`,
  );
};

async function parse(file) {
  const wb = XLSX.read(await file.arrayBuffer());
  const [h, ...r] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: "",
  });
  const heads = h
    .slice(3)
    .map((x) => String(x).trim())
    .filter(Boolean);
  const rows = r
    .filter((x) => String(x[0]).trim() || String(x[1]).trim())
    .map((x) => [
      String(x[0]).trim(),
      String(x[1]).trim(),
      String(x[2]).trim(),
      ...heads.map((_, i) => Number(x[i + 3]) || 0),
    ]);
  return { heads, rows };
}

// Firestore no admite arrays anidados: las filas se guardan como JSON en lotes de 400
async function publish(d) {
  const size = 400,
    n = Math.ceil(d.rows.length / size),
    b = writeBatch(db),
    at = Date.now();
  for (let i = 0; i < n; i++)
    b.set(doc(db, "inventory", "c" + i), {
      json: JSON.stringify(d.rows.slice(i * size, (i + 1) * size)),
    });
  b.set(doc(db, "config", "inventory"), { heads: d.heads, chunks: n, at });
  await b.commit();
  return at;
}
async function load() {
  const m = await getDoc(doc(db, "config", "inventory"));
  if (!m.exists()) return null;
  const { heads, chunks, at } = m.data(),
    parts = {};
  (await getDocs(collection(db, "inventory"))).forEach(
    (x) => (parts[x.id] = JSON.parse(x.data().json)),
  );
  return {
    heads,
    at,
    rows: Array.from({ length: chunks }, (_, i) => parts["c" + i] || []).flat(),
  };
}

// Desplegable con buscador: al escribir se reduce la lista (ignora tildes y mayúsculas; busca en el nombre mostrado y en el original)
function Combo({ value, onChange, options, placeholder, disabled, title }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const box = useRef(null);
  useEffect(() => {
    if (!open) return;
    const out = (e) => {
      if (!box.current?.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", out);
    return () => document.removeEventListener("mousedown", out);
  }, [open]);
  const toks = norm(text).split(" ").filter(Boolean);
  const list = options.filter((o) =>
    toks.every((w) => norm(o.label + " " + o.value).includes(w)),
  );
  const pick = (v) => {
    onChange(v);
    setOpen(false);
  };
  const sel = options.find((o) => o.value === value);
  return (
    <div ref={box} className="relative">
      <button
        type="button"
        disabled={disabled}
        title={title}
        aria-expanded={open}
        onClick={() => {
          setOpen((o) => !o);
          setText("");
        }}
        className={
          inp +
          " flex w-64 items-center justify-between gap-2 text-left disabled:opacity-50"
        }
      >
        <span className="truncate">{sel ? sel.label : placeholder}</span>
        <ChevronDown size={14} className="shrink-0 opacity-60" />
      </button>
      {open && (
        <div className="absolute z-30 mt-1 w-72 rounded border border-slate-300 bg-white shadow-lg">
          <input
            autoFocus
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder="Escribe para buscar…"
            className="w-full rounded-t border-b border-slate-200 px-3 py-1.5 text-sm outline-none"
            onKeyDown={(e) => {
              if (e.key === "Escape") setOpen(false);
              else if (e.key === "Enter" && list[0]) pick(list[0].value);
            }}
          />
          <ul className="max-h-60 overflow-y-auto py-1 text-sm">
            {!toks.length && (
              <li>
                <button
                  type="button"
                  onClick={() => pick("")}
                  className="w-full px-3 py-1.5 text-left text-slate-500 hover:bg-slate-100"
                >
                  {placeholder}
                </button>
              </li>
            )}
            {list.map((o) => (
              <li key={o.value}>
                <button
                  type="button"
                  onClick={() => pick(o.value)}
                  className={`w-full px-3 py-1.5 text-left hover:bg-slate-100 ${o.value === value ? "font-semibold" : ""}`}
                >
                  {o.label}
                </button>
              </li>
            ))}
            {!list.length && (
              <li className="px-3 py-2 text-slate-400">Sin resultados</li>
            )}
          </ul>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(undefined);
  useEffect(() => onAuthStateChanged(auth, setUser), []);
  if (user === undefined) return null;
  return user ? <Main user={user} /> : <Login />;
}

function Login() {
  const [e, setE] = useState(""),
    [p, setP] = useState(""),
    [err, setErr] = useState("");
  const go = async (ev) => {
    ev.preventDefault();
    try {
      await signInWithEmailAndPassword(auth, e, p);
    } catch {
      setErr("Correo o contraseña incorrectos");
    }
  };
  return (
    <div className="grid min-h-screen place-items-center bg-slate-100">
      <form
        onSubmit={go}
        className="w-80 space-y-3 rounded-lg bg-white p-6 shadow"
      >
        <h1 className="text-lg font-semibold">Existencias por sucursal</h1>
        <input
          className={inp + " w-full"}
          type="email"
          placeholder="Correo"
          value={e}
          onChange={(x) => setE(x.target.value)}
          required
        />
        <input
          className={inp + " w-full"}
          type="password"
          placeholder="Contraseña"
          value={p}
          onChange={(x) => setP(x.target.value)}
          required
        />
        {err && <p className="text-sm text-red-600">{err}</p>}
        <button className="w-full rounded bg-slate-900 py-2 text-white">
          Entrar
        </button>
      </form>
    </div>
  );
}

function Main({ user }) {
  const admin = ADMINS.includes(user.email.toLowerCase());
  const [data, setData] = useState(null);
  const [tab, setTab] = useState("inv");
  const [globalPin, setGlobalPin] = useState([]);
  const [localPin, setLocalPin] = useState(undefined); // undefined = seguir al admin; lista = fijados solo en esta sesión
  const [prio, setPrio] = useState(() => Array.from({ length: 15 }, blank));
  const [dict, setDict] = useState(() => Array.from({ length: 15 }, blank2));
  const [q, setQ] = useState(""),
    [cat, setCat] = useState(""),
    [target, setTarget] = useState(""),
    [limit, setLimit] = useState(300);
  const [msg, setMsg] = useState("");
  const [maxStores, setMaxStores] = useState(0); // No exhibido: "exhibido en hasta N tiendas" (0 = solo en los almacenes fijados)
  const [sort, setSort] = useState(null); // null = sin orden; { c: columna (índice en la fila), d: 1 asc | -1 desc }

  useEffect(() => {
    load()
      .then(setData)
      .catch(() => {});
    getDoc(doc(db, "config", "priorities"))
      .then((s) => s.exists() && setPrio(toRows(JSON.parse(s.data().json))))
      .catch(() => {});
    getDoc(doc(db, "config", "dictionary"))
      .then((s) => s.exists() && setDict(JSON.parse(s.data().json)))
      .catch(() => {});
    return onSnapshot(doc(db, "config", "main"), (s) =>
      setGlobalPin(asList(s.data()?.pinned)),
    );
  }, []);

  const heads = data?.heads ?? [];
  const wanted = localPin === undefined ? globalPin : localPin;
  const pins = useMemo(
    () =>
      [
        ...new Set(
          wanted
            .map((w) => heads.find((h) => norm(h) === norm(w)))
            .filter(Boolean),
        ),
      ].slice(0, MAX_PINS),
    [wanted, heads],
  );
  const opts = heads.filter((h) => !pins.includes(h)); // sucursales elegibles (Simplificar a existencia 0 y Prioritarios): todos los almacenes menos los fijados
  const tgt = pins.length && opts.includes(target) ? target : ""; // sucursal a simplificar (solo si sigue siendo válida)
  const shown = [...pins, ...(tgt ? [tgt] : opts)];
  const col = (h) => 3 + heads.indexOf(h);
  const vis = tab === "noex" ? pins : shown; // columnas de almacén visibles en la pestaña actual
  const STORES = 3 + heads.length; // índice de la columna "Exhibido en" (solo existe en No exhibido)
  const sc =
    sort &&
    (sort.c < 3 ||
      (tab === "noex" && sort.c === STORES) ||
      vis.some((h) => col(h) === sort.c))
      ? sort.c
      : -1; // columna ordenada (solo si está visible)
  const sd = sort?.d;
  // Diccionario: nombre corto de cada localidad (se busca por el nombre original, sin tildes, mayúsculas ni asteriscos)
  const dmap = useMemo(
    () =>
      new Map(
        dict
          .filter((r) => r[0].trim() && r[1].trim())
          .map((r) => [norm(r[0]), r[1].trim()]),
      ),
    [dict],
  );
  const lbl = (h) =>
    keys(h)
      .map((k) => dmap.get(k))
      .find(Boolean) ?? clean(h);
  const names = pins.map(lbl).join(", ");
  const tiendas = maxStores === 1 ? "tienda" : "tiendas";
  const cats = useMemo(
    () =>
      [...new Set((data?.rows ?? []).map((r) => r[2]).filter(Boolean))].sort(),
    [data],
  );

  const pmap = useMemo(() => {
    const m = new Map();
    prio.forEach(([c, br, all]) => {
      const k = norm(c);
      if (!k) return;
      const e = m.get(k) ?? { all: false, br: new Set() };
      if (all) e.all = true;
      else if (br.trim()) e.br.add(norm(br));
      m.set(k, e);
    });
    return m;
  }, [prio]);

  const view = useMemo(() => {
    if (!data) return [];
    const toks = norm(q).split(" ").filter(Boolean);
    const ti = tgt ? heads.indexOf(tgt) + 3 : -1;
    const pis = pins.map((h) => heads.indexOf(h) + 3);
    const tk = tgt ? keys(tgt) : [];
    const out = [];
    for (const r of data.rows) {
      if (cat && r[2] !== cat) continue;
      if (toks.length) {
        const t = norm(r[0] + " " + r[1]);
        if (!toks.every((w) => t.includes(w))) continue;
      }
      if (ti > 0 && !(r[ti] <= 0 && pis.some((i) => r[i] >= 1))) continue; // 0 en la sucursal y existencia en alguno de los almacenes fijados
      const p = pmap.get(norm(r[0]));
      const star =
        !!p &&
        (p.all || (ti > 0 ? tk.some((k) => p.br.has(k)) : p.br.size > 0));
      out.push({ r, star });
    }
    // prioritarios primero; dentro de cada grupo, la columna elegida (orden estable)
    return out.sort(
      (a, b) => b.star - a.star || (sc < 0 ? 0 : by(a.r[sc], b.r[sc]) * sd),
    );
  }, [data, q, cat, tgt, pins, pmap, sc, sd]);

  // No exhibido: artículos con existencia en alguno de los almacenes fijados y en hasta maxStores tiendas (las demás localidades). st = tiendas con existencia
  const noex = useMemo(() => {
    if (!data || !pins.length) return [];
    const pis = pins.map((h) => heads.indexOf(h) + 3);
    const sis = opts.map(col);
    const toks = norm(q).split(" ").filter(Boolean);
    const out = [];
    for (const r of data.rows) {
      if (!pis.some((i) => r[i] > 0)) continue;
      const st = opts.filter((h, k) => r[sis[k]] > 0);
      if (st.length > maxStores) continue;
      if (cat && r[2] !== cat) continue;
      if (toks.length) {
        const t = norm(r[0] + " " + r[1]);
        if (!toks.every((w) => t.includes(w))) continue;
      }
      out.push({ r, st });
    }
    const key = (x) => (sc === STORES ? x.st.length : x.r[sc]); // "Exhibido en" se ordena por cantidad de tiendas
    return sc < 0 ? out : out.sort((a, b) => by(key(a), key(b)) * sd);
  }, [data, pins, maxStores, q, cat, sc, sd]);

  const onFile = async (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    try {
      setData(await parse(f));
      setTarget("");
      setLimit(300);
      setSort(null);
      setMsg("");
    } catch {
      setMsg("No se pudo leer el Excel");
    }
  };
  const share = async () => {
    try {
      const at = await publish(data);
      setData((d) => ({ ...d, at }));
      setMsg("Datos publicados para todos");
    } catch {
      setMsg("No se pudo publicar (revisa reglas de Firestore)");
    }
  };
  // Fijar / desfijar un almacén (hasta MAX_PINS). El admin lo guarda para todos; los demás, solo en su sesión
  const togglePin = async (h) => {
    const next = pins.includes(h)
      ? pins.filter((x) => x !== h)
      : [...pins, h].slice(0, MAX_PINS);
    setTarget("");
    if (admin) {
      await setDoc(doc(db, "config", "main"), { pinned: next });
      setLocalPin(undefined);
    } else setLocalPin(next);
  };

  // Descargar a Excel exactamente lo que se ve en la tabla (con el filtro y el orden actuales, todas las filas)
  const exportInv = () => {
    const star = view.some((x) => x.star);
    download(
      "Inventario",
      [
        "Código",
        "Descripción",
        "Categoría",
        ...shown.map(lbl),
        ...(star ? ["Prioritario"] : []),
      ],
      view.map((x) => [
        x.r[0],
        x.r[1],
        x.r[2],
        ...shown.map((h) => x.r[col(h)]),
        ...(star ? [x.star ? "Sí" : ""] : []),
      ]),
    );
  };
  const exportNoex = () =>
    download(
      "No exhibido",
      ["Código", "Descripción", "Categoría", ...pins.map(lbl), "Exhibido en"],
      noex.map(({ r, st }) => [
        r[0],
        r[1],
        r[2],
        ...pins.map((h) => r[col(h)]),
        st.map(lbl).join(" | "),
      ]),
    );

  // Clic en un encabezado: sin orden → menor a mayor → mayor a menor → sin orden
  const toggleSort = (c) => {
    setSort((s) =>
      s?.c !== c ? { c, d: 1 } : s.d === 1 ? { c, d: -1 } : null,
    );
    setLimit(300);
  };
  const sortHead = (c, label, cls = "") => {
    const on = sc === c,
      Icon = !on ? ArrowUpDown : sd === 1 ? ArrowUp : ArrowDown;
    const tip = !on
      ? "Ordenar de menor a mayor"
      : sd === 1
        ? "Ordenar de mayor a menor"
        : "Quitar orden";
    return (
      <button
        onClick={() => toggleSort(c)}
        title={tip}
        className="flex items-start gap-1 rounded text-left hover:bg-black/10"
      >
        <span className={cls}>{label}</span>
        <Icon
          size={16}
          className={`mt-px shrink-0 ${on ? "" : "opacity-40"}`}
        />
      </button>
    );
  };

  const setCell = (i, j, v) =>
    setPrio((p) =>
      p.map((r, k) => (k === i ? r.map((c, m) => (m === j ? v : c)) : r)),
    );
  const onPaste = (e, i, j) => {
    const t = e.clipboardData.getData("text");
    if (!/[\t\n]/.test(t.trim())) return;
    e.preventDefault();
    const grid = t
      .replace(/\r/g, "")
      .replace(/\n+$/, "")
      .split("\n")
      .map((l) => l.split("\t"));
    setPrio((p) => {
      const n = p.map((r) => [...r]);
      grid.forEach((g, a) => {
        while (n.length <= i + a) n.push(blank());
        g.forEach((c, b) => {
          if (j + b < 3)
            n[i + a][j + b] = j + b === 2 ? YES.test(c.trim()) : c.trim();
        });
      });
      return n;
    });
  };
  const savePrio = async () => {
    try {
      await setDoc(doc(db, "config", "priorities"), {
        json: JSON.stringify(prio.filter((r) => r[0].trim())),
      });
      setMsg("Prioritarios guardados para todos");
    } catch {
      setMsg("No se pudo guardar");
    }
  };

  // Diccionario de nombres de localidades: [nombre original, nombre corto]
  const setD = (i, j, v) =>
    setDict((p) =>
      p.map((r, k) => (k === i ? r.map((c, m) => (m === j ? v : c)) : r)),
    );
  const pasteD = (e, i, j) => {
    const t = e.clipboardData.getData("text");
    if (!/[\t\n]/.test(t.trim())) return;
    e.preventDefault();
    const grid = t
      .replace(/\r/g, "")
      .replace(/\n+$/, "")
      .split("\n")
      .map((l) => l.split("\t"));
    setDict((p) => {
      const n = p.map((r) => [...r]);
      grid.forEach((g, a) => {
        while (n.length <= i + a) n.push(blank2());
        g.forEach((c, b) => {
          if (j + b < 2) n[i + a][j + b] = c.trim();
        });
      });
      return n;
    });
  };
  // Agrega al diccionario los almacenes del inventario que todavía no están (con el nombre corto vacío para llenarlo)
  const addHeads = () =>
    setDict((p) => {
      const rows = p.filter((r) => r[0].trim() || r[1].trim());
      const have = new Set(rows.map((r) => norm(r[0])));
      return [
        ...rows,
        ...heads
          .filter((h) => !keys(h).some((k) => have.has(k)))
          .map((h) => [clean(h), ""]),
      ];
    });
  const saveDict = async () => {
    try {
      await setDoc(doc(db, "config", "dictionary"), {
        json: JSON.stringify(dict.filter((r) => r[0].trim())),
      });
      setMsg("Diccionario guardado para todos");
    } catch {
      setMsg("No se pudo guardar");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <header className="flex flex-wrap items-center gap-3 bg-slate-900 px-4 py-3 text-white">
        <h1 className="mr-auto text-lg font-semibold">
          Existencias por sucursal
        </h1>
        <label className="flex cursor-pointer items-center gap-2 rounded bg-white/10 px-3 py-1.5 text-sm hover:bg-white/20">
          <Upload size={16} /> Cargar Excel
          <input type="file" accept=".xlsx,.xls" hidden onChange={onFile} />
        </label>
        {admin && data && (
          <button
            onClick={share}
            className="flex items-center gap-2 rounded bg-emerald-600 px-3 py-1.5 text-sm hover:bg-emerald-500"
          >
            <Share2 size={16} /> Publicar para todos
          </button>
        )}
        <span className="text-xs opacity-70">
          {user.email}
          {admin && " (admin)"}
        </span>
        <button onClick={() => signOut(auth)} title="Salir">
          <LogOut size={16} />
        </button>
      </header>

      <nav className="flex gap-1 border-b bg-white px-4">
        {[
          ["inv", "Inventario"],
          ["prio", "Prioritarios"],
          ["noex", "No exhibido"],
          ["dic", "Diccionario"],
        ].map(([k, l]) => (
          <button
            key={k}
            onClick={() => {
              setTab(k);
              setLimit(300);
              if (k === "noex") setMaxStores(0);
            }}
            className={`px-4 py-2 text-sm ${tab === k ? "border-b-2 border-slate-900 font-semibold" : "text-slate-500"}`}
          >
            {l}
          </button>
        ))}
        {msg && (
          <span className="ml-auto self-center text-sm text-emerald-700">
            {msg}
          </span>
        )}
      </nav>

      {tab === "inv" ? (
        <main className="space-y-3 p-4">
          {!data ? (
            <p className="py-16 text-center text-slate-500">
              Carga un Excel para empezar (código, descripción, categoría y un
              almacén por columna).
            </p>
          ) : (
            <>
              <div className="flex flex-wrap items-center gap-2">
                <input
                  className={inp + " w-64"}
                  placeholder="Buscar código o descripción…"
                  value={q}
                  onChange={(e) => {
                    setQ(e.target.value);
                    setLimit(300);
                  }}
                />
                <select
                  className={inp}
                  value={cat}
                  onChange={(e) => {
                    setCat(e.target.value);
                    setLimit(300);
                  }}
                >
                  <option value="">Todas las categorías</option>
                  {cats.map((c) => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
                <Combo
                  value={tgt}
                  onChange={(v) => {
                    setTarget(v);
                    setLimit(300);
                  }}
                  options={opts.map((h) => ({ value: h, label: lbl(h) }))}
                  placeholder="Seleccione tienda"
                  disabled={!pins.length}
                  title={pins.length ? "" : "Fija un almacén primero"}
                />
                <span className="text-sm text-slate-500">
                  {view.length} ítems{pins.length > 0 && ` · fijados: ${names}`}
                </span>
                <button
                  onClick={exportInv}
                  disabled={!view.length}
                  className="flex items-center gap-1 rounded border bg-white px-3 py-1.5 text-sm hover:bg-slate-100 disabled:opacity-50"
                >
                  <Download size={14} /> Descargar Excel
                </button>
                <span className="ml-auto text-sm text-slate-500">
                  {Number.isFinite(data.at)
                    ? `Publicado: ${fdate.format(data.at)}`
                    : "Sin publicar (archivo local)"}
                </span>
              </div>
              <div className="max-h-[75vh] overflow-auto rounded border bg-white">
                <table className="w-full border-collapse text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-200">
                    <tr>
                      {["Código", "Descripción", "Categoría"].map((t, i) => (
                        <th
                          key={t}
                          className="px-2 py-2 text-left align-bottom"
                        >
                          {sortHead(i, t)}
                        </th>
                      ))}
                      {shown.map((h) => (
                        <th
                          key={h}
                          className={`min-w-32 px-2 py-2 text-left align-bottom ${pins.includes(h) ? "bg-sky-200" : ""}`}
                        >
                          <div className="flex items-start gap-1">
                            {sortHead(
                              col(h),
                              lbl(h),
                              "text-xs font-semibold leading-tight",
                            )}
                            <button
                              onClick={() => togglePin(h)}
                              disabled={
                                !pins.includes(h) && pins.length >= MAX_PINS
                              }
                              title={
                                pins.includes(h)
                                  ? "Desfijar"
                                  : pins.length >= MAX_PINS
                                    ? `Máximo ${MAX_PINS} almacenes fijados`
                                    : "Fijar"
                              }
                              className="shrink-0 rounded p-0.5 hover:bg-black/10 disabled:opacity-30"
                            >
                              {pins.includes(h) ? (
                                <PinOff size={14} />
                              ) : (
                                <Pin size={14} />
                              )}
                            </button>
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {view.slice(0, limit).map(({ r, star }, i) => (
                      <tr
                        key={i}
                        className={
                          star
                            ? "bg-amber-100"
                            : "odd:bg-white even:bg-slate-50"
                        }
                      >
                        <td className="px-2 py-1 font-mono">{r[0]}</td>
                        <td className="px-2 py-1">
                          {r[1]}
                          {star && (
                            <span className="ml-2 rounded bg-amber-500 px-1.5 py-0.5 text-[10px] font-semibold text-white">
                              PRIORITARIO
                            </span>
                          )}
                        </td>
                        <td className="px-2 py-1 text-slate-500">{r[2]}</td>
                        {shown.map((h) => {
                          const v = r[col(h)];
                          return (
                            <td
                              key={h}
                              className={`px-2 py-1 text-center tabular-nums ${pins.includes(h) ? "bg-sky-50 font-semibold" : ""} ${v <= 0 ? "text-red-500" : ""}`}
                            >
                              {v}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {view.length > limit && (
                <button
                  onClick={() => setLimit((l) => l + 300)}
                  className="rounded border bg-white px-3 py-1.5 text-sm"
                >
                  Mostrar más ({view.length - limit} restantes)
                </button>
              )}
            </>
          )}
        </main>
      ) : tab === "prio" ? (
        <main className="space-y-3 p-4">
          <p className="text-sm text-slate-500">
            Pega desde Excel (Ctrl+V) en cualquier celda. Para varias sucursales
            del mismo código, agrega una fila por sucursal.{" "}
            {!admin && "Tus cambios solo aplican en tu sesión."}
          </p>
          <div className="overflow-auto rounded border bg-white">
            <table className="w-full border-collapse text-sm">
              <thead className="bg-slate-200">
                <tr>
                  {[
                    "Código",
                    "Sucursal",
                    "Prioritario para todas las sucursales",
                  ].map((t) => (
                    <th key={t} className="px-2 py-2 text-left">
                      {t}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {prio.map((r, i) => {
                  // Sucursal tal como figura en el inventario (si no coincide con ninguna, se muestra el texto guardado)
                  const cur =
                    (r[1] && opts.find((h) => keys(h).includes(norm(r[1])))) ||
                    r[1];
                  return (
                    <tr key={i} className="border-t">
                      <td className="p-1">
                        <input
                          className="w-full rounded border border-slate-300 px-2 py-1"
                          value={r[0]}
                          onChange={(e) => setCell(i, 0, e.target.value)}
                          onPaste={(e) => onPaste(e, i, 0)}
                        />
                      </td>
                      <td className="p-1">
                        <select
                          className="w-full rounded border border-slate-300 bg-white px-2 py-1 disabled:bg-slate-100"
                          value={cur}
                          disabled={r[2]}
                          onChange={(e) => setCell(i, 1, e.target.value)}
                          onPaste={(e) => onPaste(e, i, 1)}
                        >
                          <option value="">Seleccionar sucursal…</option>
                          {cur && !opts.includes(cur) && (
                            <option value={cur}>{clean(cur)}</option>
                          )}
                          {opts.map((h) => (
                            <option key={h} value={h}>
                              {lbl(h)}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td className="p-1 text-center">
                        <input
                          type="checkbox"
                          checked={r[2]}
                          onChange={(e) => setCell(i, 2, e.target.checked)}
                          onPaste={(e) => onPaste(e, i, 2)}
                        />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="flex gap-2">
            <button
              onClick={() =>
                setPrio((p) => [...p, ...Array.from({ length: 10 }, blank)])
              }
              className="flex items-center gap-1 rounded border bg-white px-3 py-1.5 text-sm"
            >
              <Plus size={14} /> 10 filas
            </button>
            {admin && (
              <button
                onClick={savePrio}
                className="rounded bg-slate-900 px-3 py-1.5 text-sm text-white"
              >
                Guardar para todos
              </button>
            )}
          </div>
        </main>
      ) : tab === "noex" ? (
        <main className="space-y-3 p-4">
          {!data ? (
            <p className="py-16 text-center text-slate-500">
              Carga un Excel para empezar (código, descripción, categoría y un
              almacén por columna).
            </p>
          ) : !pins.length ? (
            <p className="rounded border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-800">
              Fija primero el almacén de distribución (puedes fijar hasta{" "}
              {MAX_PINS}) para ver aquí los artículos que solo tienen existencia
              en él: ve a la pestaña Inventario y pulsa el icono de pin en el
              encabezado de ese almacén.{" "}
              <button
                onClick={() => setTab("inv")}
                className="font-semibold underline"
              >
                Ir a Inventario
              </button>
            </p>
          ) : (
            <>
              <div className="flex flex-wrap items-center gap-2">
                <input
                  className={inp + " w-64"}
                  placeholder="Buscar código o descripción…"
                  value={q}
                  onChange={(e) => {
                    setQ(e.target.value);
                    setLimit(300);
                  }}
                />
                <select
                  className={inp}
                  value={cat}
                  onChange={(e) => {
                    setCat(e.target.value);
                    setLimit(300);
                  }}
                >
                  <option value="">Todas las categorías</option>
                  {cats.map((c) => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
                <label className="flex items-center gap-2 text-sm text-slate-600">
                  Exhibido en hasta:
                  <select
                    className={inp}
                    value={maxStores}
                    onChange={(e) => {
                      setMaxStores(Number(e.target.value));
                      setLimit(300);
                    }}
                  >
                    {Array.from({ length: MAX_STORES + 1 }, (_, n) => (
                      <option key={n} value={n}>
                        {n} {n === 1 ? "tienda" : "tiendas"}
                      </option>
                    ))}
                  </select>
                </label>
                <span className="text-sm text-slate-500">
                  {noex.length} ítems con existencia{" "}
                  {maxStores ? "en" : "solo en"} {names}
                  {maxStores > 0 &&
                    ` y exhibidos en hasta ${maxStores} ${tiendas}`}
                </span>
                <button
                  onClick={exportNoex}
                  disabled={!noex.length}
                  className="flex items-center gap-1 rounded border bg-white px-3 py-1.5 text-sm hover:bg-slate-100 disabled:opacity-50"
                >
                  <Download size={14} /> Descargar Excel
                </button>
              </div>
              {noex.length === 0 ? (
                <p className="py-16 text-center text-slate-500">
                  {maxStores
                    ? `Ningún ítem con existencia en ${names} está exhibido en hasta ${maxStores} ${tiendas}`
                    : `Ningún ítem tiene existencia únicamente en ${names}`}
                  {(q || cat) && " con esos filtros"}.
                </p>
              ) : (
                <div className="max-h-[75vh] overflow-auto rounded border bg-white">
                  <table className="w-full border-collapse text-sm">
                    <thead className="sticky top-0 z-10 bg-slate-200">
                      <tr>
                        {["Código", "Descripción", "Categoría"].map((t, i) => (
                          <th
                            key={t}
                            className="px-2 py-2 text-left align-bottom"
                          >
                            {sortHead(i, t)}
                          </th>
                        ))}
                        {pins.map((h) => (
                          <th
                            key={h}
                            className="min-w-32 bg-sky-200 px-2 py-2 text-left align-bottom"
                          >
                            {sortHead(
                              col(h),
                              lbl(h),
                              "text-xs font-semibold leading-tight",
                            )}
                          </th>
                        ))}
                        <th className="min-w-40 px-2 py-2 text-left align-bottom">
                          {sortHead(STORES, "Exhibido en")}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {noex.slice(0, limit).map(({ r, st }, i) => (
                        <tr key={i} className="odd:bg-white even:bg-slate-50">
                          <td className="px-2 py-1 font-mono">{r[0]}</td>
                          <td className="px-2 py-1">{r[1]}</td>
                          <td className="px-2 py-1 text-slate-500">{r[2]}</td>
                          {pins.map((h) => {
                            const v = r[col(h)];
                            return (
                              <td
                                key={h}
                                className={`bg-sky-50 px-2 py-1 text-center font-semibold tabular-nums ${v <= 0 ? "text-red-500" : ""}`}
                              >
                                {v}
                              </td>
                            );
                          })}
                          <td className="px-2 py-1 text-slate-600">
                            {st.map(lbl).join(" | ")}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              {noex.length > limit && (
                <button
                  onClick={() => setLimit((l) => l + 300)}
                  className="rounded border bg-white px-3 py-1.5 text-sm"
                >
                  Mostrar más ({noex.length - limit} restantes)
                </button>
              )}
            </>
          )}
        </main>
      ) : (
        <main className="space-y-3 p-4">
          <p className="text-sm text-slate-500">
            Escribe (o pega desde Excel con Ctrl+V) el nombre completo de la
            localidad y el nombre corto con el que quieres verla en las tablas.{" "}
            {!admin && "Tus cambios solo aplican en tu sesión."}
          </p>
          <div className="overflow-auto rounded border bg-white">
            <table className="w-full border-collapse text-sm">
              <thead className="bg-slate-200">
                <tr>
                  {[
                    "Nombre original (como viene en el Excel)",
                    "Nombre corto",
                  ].map((t) => (
                    <th key={t} className="px-2 py-2 text-left">
                      {t}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {dict.map((r, i) => (
                  <tr key={i} className="border-t">
                    {[0, 1].map((j) => (
                      <td key={j} className="p-1">
                        <input
                          className="w-full rounded border border-slate-300 px-2 py-1"
                          value={r[j]}
                          onChange={(e) => setD(i, j, e.target.value)}
                          onPaste={(e) => pasteD(e, i, j)}
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              onClick={() =>
                setDict((p) => [...p, ...Array.from({ length: 10 }, blank2)])
              }
              className="flex items-center gap-1 rounded border bg-white px-3 py-1.5 text-sm"
            >
              <Plus size={14} /> 10 filas
            </button>
            {heads.length > 0 && (
              <button
                onClick={addHeads}
                className="rounded border bg-white px-3 py-1.5 text-sm"
              >
                Agregar los almacenes del inventario
              </button>
            )}
            {admin && (
              <button
                onClick={saveDict}
                className="rounded bg-slate-900 px-3 py-1.5 text-sm text-white"
              >
                Guardar para todos
              </button>
            )}
          </div>
        </main>
      )}
    </div>
  );
}
