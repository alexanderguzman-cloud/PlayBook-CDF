// Buscador del formulario de la alianza MYWAK por teléfono del cliente. Solo con sesión iniciada; la consulta la hace el Apps Script.
import { leer, json } from '../comun/session.ts';
export default async (req: Request) => {
  if (req.method !== 'POST') return json({ ok: false }, 405);
  const s = await leer(req); if (!s) return json({ ok: false, error: 'sesion' }, 401);
  const url = Netlify.env.get('SHEETS_URL'), token = Netlify.env.get('SHEETS_TOKEN');
  if (!url || !token) return json({ ok: false, error: 'Falta configurar SHEETS_URL o SHEETS_TOKEN' }, 500);
  const b = await req.json().catch(() => ({}));
  const telefono = String((b as any).telefono ?? '').replace(/\D/g, '').slice(-15);
  if (telefono.length < 7) return json({ ok: false, error: 'telefono' }, 400);
  try {
    const r = await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ tipo: 'mywak', token, email: s.email, telefono }), redirect: 'follow' });
    const j = await r.json().catch(() => ({ ok: false }));
    if (!j || !(j as any).ok) return json({ ok: false }, 502);
    if (!Array.isArray((j as any).filas)) return json({ ok: false, error: 'sin_buscador' }, 502); // el Apps Script publicado aún no tiene la búsqueda MYWAK
    return json({ ok: true, filas: (j as any).filas });
  } catch { return json({ ok: false }, 502); }
};
export const config = { path: '/api/mywak' };
