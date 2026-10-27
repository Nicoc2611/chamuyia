# ChamuyIA

Tu fichero de chamuyo: subís la captura del chat y te arma tres opciones con tu forma de escribir. También guarda a cada chica con su historial, da ideas de citas y revisa tus fotos y tu perfil.

## Cómo está armada

- **La app** (`index.html`) se publica gratis con GitHub Pages.
- **El servidor** (`worker.js`) es un Worker de Cloudflare. Guarda la API key de Claude, pide una clave de acceso y sincroniza el fichero entre el celu y la compu (Workers KV).

La API key nunca llega al navegador. Las chicas, los chats y el perfil se guardan en Cloudflare, no en este repo.

## Configuración del Worker

| Nombre | Tipo | Para qué |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | Secreto | API key de [console.anthropic.com](https://console.anthropic.com) |
| `CLAVE` | Secreto | La clave que se pone en la app para entrar |
| `DATOS` | Binding de KV | Donde se guarda el fichero |
| `ORIGEN` | Variable (opcional) | Desde dónde se abre la app. Por defecto `https://nicoc2611.github.io` |
| `MODELO`, `MODELO_RAPIDO` | Variables (opcionales) | Para cambiar de modelo. Por defecto `claude-sonnet-5-5` y `claude-haiku-5-5` |

La dirección del Worker (`https://chamuyia.nicoc2611.workers.dev`) va en la constante `SERVIDOR` de `index.html`.

## Copias

Desde el menú de tu inicial: **Descargar copia de mis datos** y **Cargar una copia**.
