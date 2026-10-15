export default async (req: Request) => new Response(null, { status: 302, headers: { location: '/login.html', 'set-cookie': 'cx_s=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0', 'cache-control': 'no-store' } });
export const config = { path: '/api/logout' };
