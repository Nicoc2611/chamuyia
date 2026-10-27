// Servidor de ChamuyIA: un Worker de Cloudflare.
// Guarda tu API key de Claude (la app nunca la ve), pide la clave de acceso
// y guarda tu fichero para que lo veas en el celu y en la compu.
//
// Variables del Worker (Settings > Variables and Secrets):
//   ANTHROPIC_API_KEY  (secreto)  tu API key de console.anthropic.com
//   CLAVE              (secreto)  la clave de acceso que vas a poner en la app
//   ORIGEN             (opcional) desde dónde se abre la app. Por defecto https://nicoc2611.github.io
//   MODELO, MODELO_RAPIDO (opcional) para cambiar de modelo sin tocar la app
// Binding (Settings > Bindings): un KV namespace con el nombre DATOS.

const ANTHROPIC = 'https://api.anthropic.com/v1/messages';
const ORIGEN = 'https://nicoc2611.github.io';

export default {
  async fetch(request, env) {
    const origen = env.ORIGEN || ORIGEN;
    const cors = {
      'Access-Control-Allow-Origin': origen,
      'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, X-Clave',
      'Access-Control-Max-Age': '86400',
      Vary: 'Origin',
    };
    const json = (obj, status = 200) =>
      new Response(JSON.stringify(obj), { status, headers: { ...cors, 'Content-Type': 'application/json' } });

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });

    const url = new URL(request.url);
    if (url.pathname === '/') return json({ ok: true, app: 'ChamuyIA' });

    if (!env.CLAVE) return json({ code: 'sin_clave', mensaje: 'Falta la variable CLAVE en el Worker' }, 500);
    if (!iguales(request.headers.get('X-Clave') || '', env.CLAVE)) return json({ code: 'clave' }, 401);

    if (url.pathname === '/ping') {
      return json({
        ok: true,
        datos: !!env.DATOS,
        claude: !!env.ANTHROPIC_API_KEY,
        modelo: env.MODELO || 'claude-sonnet-5-5',
        modeloRapido: env.MODELO_RAPIDO || 'claude-haiku-5-5',
      });
    }

    if (url.pathname === '/claude' && request.method === 'POST') {
      if (!env.ANTHROPIC_API_KEY) return json({ code: 'api_key' }, 500);
      const r = await fetch(ANTHROPIC, {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
          'x-api-key': env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
        },
        body: await request.arrayBuffer(),
      });
      return new Response(r.body, { status: r.status, headers: { ...cors, 'Content-Type': 'application/json' } });
    }

    if (url.pathname === '/datos') {
      if (!env.DATOS) return json({ code: 'sin_datos' }, 501);
      if (request.method === 'GET') {
        const v = await env.DATOS.get('fichero');
        return new Response(v || 'null', { headers: { ...cors, 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });
      }
      if (request.method === 'PUT') {
        const cuerpo = await request.text();
        if (!cuerpo.startsWith('{')) return json({ code: 'formato' }, 400);
        await env.DATOS.put('fichero', cuerpo);
        return json({ ok: true });
      }
    }

    return json({ code: 'no_existe' }, 404);
  },
};

// Compara la clave sin cortar en el primer caracter distinto.
function iguales(a, b) {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
