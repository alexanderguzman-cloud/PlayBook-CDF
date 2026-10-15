// Clics, comentarios y preguntas sin respuesta del asistente hacia Google Sheets (Apps Script). El correo sale de la sesión, nunca del navegador.
import { leer, json } from '../comun/session.ts';
const corta = (v: unknown, n: number) => String(v ?? '').slice(0, n);
export default async (req: Request) => {
  if (req.method !== 'POST') return json({ ok: false }, 405);
  const s = await leer(req); if (!s) return json({ ok: false }, 401);
  const url = Netlify.env.get('SHEETS_URL'), token = Netlify.env.get('SHEETS_TOKEN');
  if (!url || !token) return json({ ok: false, error: 'Falta configurar SHEETS_URL o SHEETS_TOKEN' }, 500);
  const b = await req.json().catch(() => ({}));
  const ruta = new URL(req.url).pathname, esFeedback = ruta.endsWith('/feedback');
  const datos = ruta.endsWith('/asistente')
    ? { tipo: 'asistente', token, email: s.email, nombre: s.nombre, pregunta: corta(b.pregunta, 500), vista: corta(b.vista, 60) }
    : esFeedback
    ? { tipo: 'feedback', token, email: s.email, nombre: s.nombre, proceso: corta(b.proceso, 200), paso: corta(b.paso, 200), clase: corta(b.clase, 60), comentario: corta(b.comentario, 3000), pagina: corta(b.pagina, 300) }
    : { tipo: 'clics', token, email: s.email, nombre: s.nombre, clics: Math.max(0, Math.min(5000, Number(b.clics) || 0)), procesos: (Array.isArray(b.procesos) ? b.procesos : []).slice(0, 60).map((x: unknown) => corta(x, 80)), busquedas: (Array.isArray(b.busquedas) ? b.busquedas : []).slice(0, 30).map((x: unknown) => corta(x, 60)).filter((x: string) => !/\d{5,}|@/.test(x)),
        eventos: (Array.isArray(b.eventos) ? b.eventos : []).slice(0, 200).map((e: any) => ({ ts: Number(e && e.ts) || Date.now(), modulo: corta(e && e.modulo, 60), proceso: corta(e && e.proceso, 120), paso: corta(e && e.paso, 160), accion: corta(e && e.accion, 160), busqueda: corta(e && e.busqueda, 120) })) };
  try {
    const r = await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(datos), redirect: 'follow' });
    return json({ ok: r.ok }, r.ok ? 200 : 502);
  } catch { return json({ ok: false }, 502); }
};
export const config = { path: ['/api/evento', '/api/feedback', '/api/asistente'] };
