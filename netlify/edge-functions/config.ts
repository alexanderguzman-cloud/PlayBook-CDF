import { json, dominio } from '../comun/session.ts';
export default async () => json({ clientId: Netlify.env.get('GOOGLE_CLIENT_ID') || '', dominio: dominio() });
export const config = { path: '/api/config' };
