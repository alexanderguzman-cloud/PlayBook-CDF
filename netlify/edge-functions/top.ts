// Lo más buscado y lo más visto por el equipo (últimos 30 días), leído del Sheet de clics a través del Apps Script.
import { leer, json } from '../comun/session.ts';
let cache: { t: number; d: unknown } | null = null;
export default async (req: Request) => {
  const s = await leer(req); if (!s) return json({ ok: false }, 401);
  if (cache && Date.now() - cache.t < 30 * 60 * 1000) return json(cache.d);
  const url = Netlify.env.get('SHEETS_URL'), token = Netlify.env.get('SHEETS_TOKEN');
  if (!url || !token) return json({ ok: false, procesos: [], busquedas: [] });
  try {
    const r = await fetch(url + '?tipo=top&token=' + encodeURIComponent(token), { redirect: 'follow' });
    const d = await r.json();
    const out = { ok: !!d.ok, procesos: Array.isArray(d.procesos) ? d.procesos.slice(0, 8) : [], busquedas: Array.isArray(d.busquedas) ? d.busquedas.slice(0, 10) : [] };
    if (out.ok) cache = { t: Date.now(), d: out };
    return json(out);
  } catch { return json({ ok: false, procesos: [], busquedas: [] }); }
};
export const config = { path: '/api/top' };
