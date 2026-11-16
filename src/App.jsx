import { useEffect, useMemo, useState } from "react";
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
const YES = /^(x|si|sí|true|1|verdadero|prioritario)$/i;
const inp = "rounded border border-slate-300 bg-white px-3 py-1.5 text-sm";
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
  const [globalPin, setGlobalPin] = useState(null);
  const [localPin, setLocalPin] = useState(undefined); // undefined = seguir al admin; null = desfijado localmente
  const [prio, setPrio] = useState(() => Array.from({ length: 15 }, blank));
  const [q, setQ] = useState(""),
    [cat, setCat] = useState(""),
    [target, setTarget] = useState(""),
    [limit, setLimit] = useState(300);
  const [msg, setMsg] = useState("");
  const [sort, setSort] = useState(null); // null = sin orden; { c: columna (índice en la fila), d: 1 asc | -1 desc }

  useEffect(() => {
    load()
      .then(setData)
      .catch(() => {});
    getDoc(doc(db, "config", "priorities"))
      .then((s) => s.exists() && setPrio(toRows(JSON.parse(s.data().json))))
      .catch(() => {});
    return onSnapshot(doc(db, "config", "main"), (s) =>
      setGlobalPin(s.data()?.pinned ?? null),
    );
  }, []);

  const heads = data?.heads ?? [];
  const wanted = localPin === undefined ? globalPin : localPin;
  const pin = wanted
    ? (heads.find((h) => norm(h) === norm(wanted)) ?? null)
    : null;
  const shown =
    pin && target
      ? [pin, target]
      : pin
        ? [pin, ...heads.filter((h) => h !== pin)]
        : heads;
  const col = (h) => 3 + heads.indexOf(h);
  const opts = heads.filter((h) => h !== pin); // sucursales elegibles (Simplificar a existencia 0 y Prioritarios): todos los almacenes menos el fijado
  const vis = tab === "noex" ? (pin ? [pin] : []) : shown; // columnas de almacén visibles en la pestaña actual
  const sc =
    sort && (sort.c < 3 || vis.some((h) => col(h) === sort.c)) ? sort.c : -1; // columna ordenada (solo si está visible)
  const sd = sort?.d;
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
    const ti = target ? heads.indexOf(target) + 3 : -1,
      pi = pin ? heads.indexOf(pin) + 3 : -1;
    const tk = target ? keys(target) : [];
    const out = [];
    for (const r of data.rows) {
      if (cat && r[2] !== cat) continue;
      if (toks.length) {
        const t = norm(r[0] + " " + r[1]);
        if (!toks.every((w) => t.includes(w))) continue;
      }
      if (ti > 0 && !(r[ti] <= 0 && r[pi] >= 1)) continue;
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
  }, [data, q, cat, target, pin, pmap, sc, sd]);

  // No exhibido: artículos con existencia solo en el almacén fijado (en todos los demás hay 0 o menos)
  const noex = useMemo(() => {
    if (!data || !pin) return [];
    const pi = heads.indexOf(pin) + 3;
    const toks = norm(q).split(" ").filter(Boolean);
    const out = data.rows.filter((r) => {
      if (!(r[pi] > 0) || r.some((v, i) => i >= 3 && i !== pi && v > 0))
        return false;
      if (cat && r[2] !== cat) return false;
      if (toks.length) {
        const t = norm(r[0] + " " + r[1]);
        if (!toks.every((w) => t.includes(w))) return false;
      }
      return true;
    });
    return sc < 0 ? out : out.sort((a, b) => by(a[sc], b[sc]) * sd);
  }, [data, pin, q, cat, sc, sd]);

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
  const togglePin = async (h) => {
    const next = pin === h ? null : h;
    setTarget("");
    if (admin) {
      await setDoc(doc(db, "config", "main"), { pinned: next });
      setLocalPin(undefined);
    } else setLocalPin(next);
  };

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
        ].map(([k, l]) => (
          <button
            key={k}
            onClick={() => {
              setTab(k);
              setLimit(300);
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
                <select
                  className={inp + " disabled:opacity-50"}
                  disabled={!pin}
                  value={target}
                  title={pin ? "" : "Fija un almacén primero"}
                  onChange={(e) => {
                    setTarget(e.target.value);
                    setLimit(300);
                  }}
                >
                  <option value="">Simplificar a existencia 0…</option>
                  {opts.map((h) => (
                    <option key={h} value={h}>
                      {clean(h)}
                    </option>
                  ))}
                </select>
                <span className="text-sm text-slate-500">
                  {view.length} ítems{pin && ` · fijado: ${clean(pin)}`}
                </span>
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
                          className={`min-w-32 px-2 py-2 text-left align-bottom ${h === pin ? "bg-sky-200" : ""}`}
                        >
                          <div className="flex items-start gap-1">
                            {sortHead(
                              col(h),
                              clean(h),
                              "text-xs font-semibold leading-tight",
                            )}
                            <button
                              onClick={() => togglePin(h)}
                              title={h === pin ? "Desfijar" : "Fijar"}
                              className="shrink-0 rounded p-0.5 hover:bg-black/10"
                            >
                              {h === pin ? (
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
                              className={`px-2 py-1 text-center tabular-nums ${h === pin ? "bg-sky-50 font-semibold" : ""} ${v <= 0 ? "text-red-500" : ""}`}
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
                              {clean(h)}
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
      ) : (
        <main className="space-y-3 p-4">
          {!data ? (
            <p className="py-16 text-center text-slate-500">
              Carga un Excel para empezar (código, descripción, categoría y un
              almacén por columna).
            </p>
          ) : !pin ? (
            <p className="rounded border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-800">
              Fija primero el almacén de distribución para ver aquí los
              artículos que solo tienen existencia en él: ve a la pestaña
              Inventario y pulsa el icono de pin en el encabezado de ese
              almacén.{" "}
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
                <span className="text-sm text-slate-500">
                  {noex.length} ítems con existencia solo en {clean(pin)}
                </span>
              </div>
              {noex.length === 0 ? (
                <p className="py-16 text-center text-slate-500">
                  Ningún ítem tiene existencia únicamente en {clean(pin)}
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
                        <th className="min-w-32 bg-sky-200 px-2 py-2 text-left align-bottom">
                          {sortHead(
                            col(pin),
                            clean(pin),
                            "text-xs font-semibold leading-tight",
                          )}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {noex.slice(0, limit).map((r, i) => (
                        <tr key={i} className="odd:bg-white even:bg-slate-50">
                          <td className="px-2 py-1 font-mono">{r[0]}</td>
                          <td className="px-2 py-1">{r[1]}</td>
                          <td className="px-2 py-1 text-slate-500">{r[2]}</td>
                          <td className="bg-sky-50 px-2 py-1 text-center font-semibold tabular-nums">
                            {r[col(pin)]}
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
      )}
    </div>
  );
}
