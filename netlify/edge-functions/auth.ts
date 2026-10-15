// Protege todo el sitio: sin sesión válida de Google @somosinternet.co se envía a /login.html
import { leer, json } from '../comun/session.ts';
export default async (req: Request, context: any) => {
  const url = new URL(req.url);
  const s = await leer(req);
  if (!s) {
    if (url.pathname.startsWith('/api/')) return json({ ok: false, error: 'Sesión vencida' }, 401);
    return Response.redirect(new URL('/login.html', url), 302);
  }
  if (url.pathname === '/' || url.pathname === '/index.html') {
    const res = await context.next();
    const html = await res.text();
    const inj = '<script>window.CX_NETLIFY=true;try{localStorage.setItem("cx-sesion",' + JSON.stringify(JSON.stringify({ email: s.email, nombre: s.nombre || '' })) + ')}catch(e){}</script>';
    const h = new Headers(res.headers); h.delete('content-length'); h.set('cache-control', 'no-store');
    return new Response(html.replace('<head>', '<head>' + inj), { status: res.status, headers: h });
  }
  return context.next();
};
export const config = { path: '/*', excludedPath: ['/login.html', '/api/login', '/api/config', '/api/logout', '/favicon.svg', '/favicon-32.png', '/apple-touch-icon.png'] };
