# Dashboard de hotelaria — Gramado e Canela

Painel estático compatível com GitHub Pages. Os dados são lidos diretamente de `dados.xlsx` sempre que alguém abre a página.

## Publicar no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie `index.html`, `styles.css`, `app.js` e `dados.xlsx` para a raiz do repositório.
3. Abra **Settings → Pages**.
4. Em **Build and deployment**, selecione **Deploy from a branch**.
5. Escolha a branch `main` e a pasta `/ (root)`, depois clique em **Save**.
6. Aguarde a publicação. O endereço será semelhante a `https://SEU-USUARIO.github.io/NOME-DO-REPOSITORIO/`.

## Atualizar os dados

Substitua `dados.xlsx` por uma versão atualizada, mantendo esse nome e a estrutura das abas. Faça o commit no GitHub. O painel carregará a nova planilha automaticamente.

As abas esperadas são:

- `Índices anuais`
- `Média de diárias`
- `Média de reservas mensais`
- `Média de taxa de ocupação diári`
- `Origem e tipo de pagamento`

Novos anos são reconhecidos automaticamente quando acrescentados à aba de índices anuais e aos cabeçalhos das abas mensais. Colunas com `estimado` no título aparecem com linha tracejada.

## Observações

- Todos os indicadores referem-se a Gramado, exceto a taxa de ocupação, que acompanha Gramado e Canela.
- Séries estimadas aparecem com linhas tracejadas e identificação explícita na legenda.
- Todas as séries de reservas são estimadas e exibidas apenas como tendência visual, sem escala numérica nem valores ao passar o cursor.
- Anos incompletos são tratados de acordo com os meses disponíveis.
- Origem e pagamento utilizam o último ano disponível se o ano selecionado não tiver esses dados.
- O painel utiliza Chart.js e SheetJS por CDN, portanto precisa de acesso à internet.
- Para testar localmente, use um servidor HTTP; abrir `index.html` diretamente pelo explorador de arquivos pode bloquear a leitura da planilha.
