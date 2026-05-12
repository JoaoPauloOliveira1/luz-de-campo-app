# Luz de Campo

Aplicativo standalone para a equipe de campo registrar novos pontos de iluminação, inclusive com uso offline preparado por RPA.

Este repositório deve ser mantido separado do Cadastro Editor. O frontend conversa com o backend público por `VITE_FIELD_API_BASE_URL`.

## Rodar Localmente

```powershell
cd "C:\Users\Everton Paiva\Desktop\Workspace\iluminPbRec\luz-de-campo-standalone"
npm install
npm run dev
```

Abra:

- http://127.0.0.1:5174

## Variável Necessária

Crie um arquivo `.env` local seguindo `.env.example`:

```env
VITE_FIELD_API_BASE_URL=https://luz-de-campo.onrender.com
```

No Render, configure a mesma variável no Static Site.

## Deploy No Render

Use este projeto como um Static Site separado.

- Build Command: `npm ci && npm run build`
- Publish Directory: `dist`
- Environment Variable: `VITE_FIELD_API_BASE_URL=https://luz-de-campo.onrender.com`

Se usar o `render.yaml`, o Render já terá esses valores como base.

## Fluxo De Uso

1. Entrar com usuário autorizado.
2. Se precisar trabalhar offline, preparar até 3 RPAs antes de sair para campo.
3. Capturar ou colar as coordenadas.
4. Confirmar o local no mapa.
5. Marcar `Implantação concluída = SIM` quando o ponto estiver pronto para ir ao Cadastro Editor.
6. Preencher os dados necessários e adicionar fotos se houver.
7. Sincronizar quando houver internet.

## Regra Importante

O Cadastro Editor só puxa para revisão os pontos sincronizados com:

- coordenada confirmada
- `Implantação concluída = SIM`
- ponto ainda não incorporado

Isso evita que rascunhos ou pontos incompletos entrem na base de luminárias por engano.
