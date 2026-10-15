// Sesión firmada (HMAC-SHA256) guardada en la cookie cx_s. La clave vive en la variable SESSION_SECRET de Netlify.
const enc = new TextEncoder();
const b64u = (bytes: Uint8Array) => btoa(String.fromCharCode(...bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const desB64u = (s: string) => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/')), c => c.charCodeAt(0));
async function clave() {
  const sec = Netlify.env.get('SESSION_SECRET') || '';
  if (sec.length < 24) throw new Error('SESSION_SECRET falta o es muy corta');
  return crypto.subtle.importKey('raw', enc.encode(sec), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
}
async function firma(p: string) { return b64u(new Uint8Array(await crypto.subtle.sign('HMAC', await clave(), enc.encode(p)))); }
export type Sesion = { email: string; nombre: string; exp: number };
export async function firmar(s: Sesion) { const p = b64u(enc.encode(JSON.stringify(s))); return p + '.' + await firma(p); }
export async function leer(req: Request): Promise<Sesion | null> {
  try {
    const m = (req.headers.get('cookie') || '').match(/(?:^|;\s*)cx_s=([^;]+)/); if (!m) return null;
    const [p, f] = m[1].split('.'); if (!p || !f || f !== await firma(p)) return null;
    const s = JSON.parse(new TextDecoder().decode(desB64u(p))) as Sesion;
    return s && s.email && s.exp > Date.now() ? s : null;
  } catch { return null; }
}
export const json = (o: unknown, status = 200, extra: Record<string, string> = {}) => new Response(JSON.stringify(o), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...extra } });
export const dominio = () => (Netlify.env.get('ALLOWED_DOMAIN') || 'somosinternet.co').toLowerCase();
