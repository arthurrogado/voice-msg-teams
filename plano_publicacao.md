# Plano de Publicação

Última atualização: 17/04/2026

## Status atual da submissão

- Versão pronta no repositório: `0.2.1`
- Pacote pronto para envio: `dist/teams-voice-message-edge-0.2.1.zip`
- Política pública em inglês: `https://telegra.ph/Privacy-Policy---Voice-Message-for-Teams-Web-04-17`
- Commit local de referência: `03376f8`
- Pendência atual: o Partner Center ainda está finalizando o cancelamento da submissão `0.2.0`; assim que esse cancelamento for concluído, a atualização `0.2.1` poderá ser publicada.

## Prazo de análise

- Depois que a submissão entrar em revisão, o próprio Partner Center informa prazo de retorno em até 7 dias úteis.

## Estrutura recomendada do repositório

- `plugin/` contém somente o que entra no zip de publicação.
- `dist/` contém o zip final pronto para submissão.
- Os arquivos na raiz continuam úteis para manutenção, documentação e regeneração de assets.

## O que entra no zip

Use apenas o conteúdo de `plugin/`:

- `manifest.json`
- `content.js`
- `styles.css`
- `popup.html`
- `popup.js`
- `popup.css`
- `mic.svg`
- `icons/`
- `_locales/`

## Como gerar o zip de submissão

```bash
cd plugin
zip -r ../dist/teams-voice-message-edge-0.2.1.zip .
```

## Idiomas

Para esta versão, faz mais sentido usar o idioma automático do navegador do que criar uma tela de configuração manual.

- Idiomas suportados: inglês e português brasileiro.
- Os textos ficam em `_locales/en/messages.json` e `_locales/pt_BR/messages.json`.
- A interface deve resolver as mensagens com `chrome.i18n`.

Uma configuração manual de idioma só passa a valer a pena se houver necessidade real de sobrescrever o idioma do navegador.

## Política de privacidade

- Mantenha a versão principal em inglês em `privacy-policy.md`.
- Mantenha a versão de apoio em português em `privacy-policy.pt-BR.md`.
- No Partner Center, use a URL pública da versão em inglês no campo `Privacy policy URL`.
- URL pública atual: `https://telegra.ph/Privacy-Policy---Voice-Message-for-Teams-Web-04-17`.

## Metadados recomendados

- Nome PT-BR: Mensagem de Voz para Teams Web
- Nome EN: Voice Message for Teams Web
- Categoria: Productivity

Descrição funcional recomendada:

Esta extensão adiciona um botão de gravação ao composer do Microsoft Teams Web, grava áudio localmente no navegador e anexa um arquivo WAV pelo fluxo nativo de arquivos do Teams. Não use a listagem para prometer o card compacto de mensagem de voz nativa do mobile, porque esse comportamento não é o que a extensão entrega hoje.

## Observação importante

Ao limpar o projeto, preserve `plugin/` como pacote publicável, mas não apague automaticamente arquivos de apoio fora dessa pasta sem confirmação explícita.