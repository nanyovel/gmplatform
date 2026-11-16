# Existencias por sucursal

1. `npm install`
2. Copia `.env.example` a `.env` y completa los datos de tu proyecto Firebase (Configuración > Tus apps > Web) y `VITE_ADMIN_EMAILS`.
3. En Firebase: activa **Authentication > Email/Password** y crea los usuarios (el admin y los demás). Crea **Firestore** y pega `firestore.rules` (con tus correos admin).
4. `npm run dev` (desarrollo) o `npm run build` (producción).

Flujo diario del admin: cargar Excel > fijar el centro de distribución (📌) > "Publicar para todos". Los demás usuarios entran y ya lo encuentran cargado y fijado (pueden fijar/desfijar localmente).
Prioritarios: se editan/pegan desde Excel en la pestaña "Prioritarios"; el admin pulsa "Guardar para todos".
Columnas del Excel: 1ª Código, 2ª Descripción, 3ª Categoría; el resto son almacenes.
