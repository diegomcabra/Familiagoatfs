# GOAT App

App de gestión (inventario + ventas + costos/PVP) hecha con React + Vite + Tailwind.
Sincroniza en vivo entre distintas terminales/dispositivos usando **Supabase** (base de datos gratis).

## 1) Crear el proyecto en Supabase (2 minutos)

1. Entrá a **https://supabase.com** y creá una cuenta gratis (podés usar GitHub o Google).
2. "New Project". Elegí un nombre y una contraseña para la base (guardala, no la vas a necesitar de nuevo para esto pero por las dudas).
3. Esperá ~1 minuto a que se cree el proyecto.
4. Andá a **SQL Editor → New query**, pegá todo el contenido del archivo `supabase.sql` (está en esta misma carpeta) y tocá **Run**. Esto crea la tabla donde se guardan tus datos.
5. Andá a **Project Settings → API**. Ahí vas a ver:
   - **Project URL** (algo como `https://xxxxx.supabase.co`)
   - **anon public key** (una clave larga)

## 2) Conectar la app a Supabase

1. En esta carpeta, hacé una copia del archivo `.env.example` y renombrala a `.env`.
2. Completá los dos valores con lo que copiaste de Supabase:
   ```
   VITE_SUPABASE_URL=https://xxxxx.supabase.co
   VITE_SUPABASE_ANON_KEY=eyJhbGciOi...
   ```
3. Guardá el archivo. **Nunca subas el `.env` a GitHub** (ya está excluido en `.gitignore`).

> Si no completás el `.env`, la app sigue funcionando igual que antes, guardando los datos solo en el navegador de esa compu (sin sincronizar). Ideal para probarla antes de conectar Supabase.

## 3) Probarla en tu compu

```bash
npm install
npm run dev
```

Abrí `http://localhost:5173`. Si el `.env` está bien configurado, vas a ver "Sincronizado" abajo a la izquierda en vez de "Guardado (solo este navegador)".

## 4) Subirla a internet con URL propia (Vercel, recomendado)

1. Subí esta carpeta a un repositorio de GitHub.
2. Entrá a **https://vercel.com**, logueate con GitHub, "Add New → Project" y elegí el repo.
3. Antes de tocar "Deploy", abrí **Environment Variables** y cargá las mismas dos variables del `.env`:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Tocá **Deploy**. En 1-2 minutos te da una URL pública (`goat-app.vercel.app`).
5. Esa misma URL, abierta desde el celular, la compu, o donde sea, va a mostrar y sincronizar los mismos datos en vivo.

### Alternativa: Netlify

Mismo esquema: `npm run build`, arrastrar la carpeta `dist` a **https://app.netlify.com/drop**, y cargar las variables de entorno en **Site settings → Environment variables** (y volver a hacer el build/deploy para que las tome).

## 5) Instalarla como app en el celular (Android)

Una vez que el sitio está publicado (Vercel o Netlify), en el celular:

1. Abrí la URL del sitio con **Chrome**.
2. Tocá el menú (los tres puntitos, arriba a la derecha).
3. Elegí **"Instalar app"** o **"Agregar a pantalla de inicio"**.
4. Confirmá. Va a aparecer un ícono de GOAT en tu pantalla de inicio, igual que cualquier app instalada.
5. Al abrirla desde ese ícono, se abre a pantalla completa (sin la barra del navegador) y carga más rápido las próximas veces.

No es necesario pasar por Google Play para esto — es el mismo sitio web, pero instalado como app (se llama "PWA"). Sigue necesitando internet para sincronizar con Supabase, igual que antes.

## Cómo funciona la sincronización

- Todos los dispositivos leen y escriben la misma fila en la tabla `app_data` de Supabase.
- Cuando alguien hace un cambio, los demás dispositivos abiertos lo reciben en vivo (sin necesidad de recargar la página), gracias a Supabase Realtime.

## Sobre la seguridad

Esta versión **no pide login**: cualquiera que tenga la URL del sitio puede ver y editar los datos. Es la configuración más simple, pensada para uso privado (compartís el link solo con quien corresponda). Si en algún momento querés sumar usuario y contraseña por persona, avisame y lo agregamos de nuevo — ya lo tenemos armado y probado.
