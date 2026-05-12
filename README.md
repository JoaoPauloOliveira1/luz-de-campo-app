# Luz de Campo

Webapp standalone voltado para a equipe de campo da rede de luminárias.

## Objetivo

- capturar a localização atual com um toque
- confirmar o local exato no mapa
- preencher o formulário de atualização do ponto
- montar uma fila local de registros
- exportar uma planilha Excel pronta para importação no sistema principal

## Rodando o projeto

```powershell
cd "C:\Users\Everton Paiva\Desktop\Workspace\iluminPbRec\luz-de-campo"
npm install
npm run dev
```

Abra:

- [http://127.0.0.1:5174](http://127.0.0.1:5174)

## Fluxo principal

1. tocar em `Capturar minha localização`
2. ajustar o marcador no mapa, se necessário
3. tocar em `Confirmar local exato`
4. preencher o formulário
5. tocar em `Adicionar registro à planilha`
6. exportar em `Exportar planilha ideal`
