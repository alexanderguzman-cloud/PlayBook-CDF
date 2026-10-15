// Valida el token de Google: cliente correcto, correo verificado y dominio @somosinternet.co. Crea la sesión (12 h).
import { firmar, json, dominio } from '../comun/session.ts';
export default async (req: Request) => {
  if (req.method !== 'POST') return json({ ok: false }, 405);
  const { credential } = await req.json().catch(() => ({ credential: '' }));
  const cid = Netlify.env.get('GOOGLE_CLIENT_ID') || '', dom = dominio();
  const r = await fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential || ''));
  if (!r.ok) return json({ ok: false, error: 'No pudimos validar tu cuenta de Google. Intenta de nuevo.' }, 401);
  const t = await r.json();
  const email = String(t.email || '').toLowerCase();
  const valido = t.aud === cid && ['accounts.google.com', 'https://accounts.google.com'].includes(t.iss) && String(t.email_verified) === 'true'
    && String(t.hd || '').toLowerCase() === dom && email.endsWith('@' + dom);
  if (!valido) return json({ ok: false, error: 'Solo pueden ingresar cuentas @' + dom + '.' }, 403);
  const v = await firmar({ email, nombre: String(t.name || ''), exp: Date.now() + 12 * 3600e3 });
  return json({ ok: true }, 200, { 'set-cookie': `cx_s=${v}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=43200` });
};
export const config = { path: '/api/login' };
