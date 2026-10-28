import { json, dominio, dominios } from '../comun/session.ts';
export default async () => json({ clientId: Netlify.env.get('GOOGLE_CLIENT_ID') || '', dominio: dominio(), dominios: dominios() });
export const config = { path: '/api/config' };
