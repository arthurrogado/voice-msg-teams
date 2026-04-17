# Plano de Publicacao

Ultima atualizacao: 17/04/2026

## Estrutura recomendada do repositorio

- `plugin/` contem somente o que entra no zip de publicacao.
- `dist/` contem o zip final pronto para submissao.
- Os arquivos na raiz continuam uteis para manutencao, documentacao e regeneracao de assets.

## O que entra no zip

Use apenas o conteudo de `plugin/`:

- `manifest.json`
- `content.js`
- `styles.css`
- `popup.html`
- `popup.js`
- `popup.css`
- `mic.svg`
- `icons/`
- `_locales/`

## Como gerar o zip de submissao

```bash
cd plugin
zip -r ../dist/teams-voice-message-edge-0.2.1.zip .
```

## Idiomas

Para esta versao, faz mais sentido usar idioma automatico do navegador do que criar uma tela de configuracao manual.

- Idiomas suportados: ingles e portugues brasileiro.
- Os textos ficam em `_locales/en/messages.json` e `_locales/pt_BR/messages.json`.
- A interface deve resolver as mensagens com `chrome.i18n`.

Configuracao manual de idioma so passa a valer a pena se houver necessidade real de sobrescrever o idioma do navegador.

## Politica de privacidade

- Mantenha a versao principal em ingles em `privacy-policy.md`.
- Mantenha a versao de apoio em portugues em `privacy-policy.pt-BR.md`.
- No Partner Center, use a URL publica da versao em ingles no campo `Privacy policy URL`.
- URL publica atual: `https://telegra.ph/Privacy-Policy---Voice-Message-for-Teams-Web-04-17`.

## Metadados recomendados

- Nome PT-BR: Mensagem de Voz para Teams Web
- Nome EN: Voice Message for Teams Web
- Categoria: Productivity

Descricao funcional recomendada:

Esta extensao adiciona um botao de gravacao ao composer do Microsoft Teams Web, grava audio localmente no navegador e anexa um arquivo WAV pelo fluxo nativo de arquivo do Teams. Nao use a listagem para prometer o card compacto de mensagem de voz nativa do mobile, porque esse comportamento nao e o que a extensao entrega hoje.

## Observacao importante

Ao limpar o projeto, preserve `plugin/` como pacote publicavel, mas nao apague automaticamente arquivos de apoio fora dessa pasta sem confirmacao explicita.