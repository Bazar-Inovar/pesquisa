# Produtos Lookup (Frontend)

Projeto frontend para carregar uma planilha de produtos e permitir busca rápida. Suporta duas fontes de dados:
- **Google Sheets API** (planilha pública ou com API key)
- **Endpoint JSON** (seu backend que retorna um array de objetos)

## Estrutura mínima do repositório
- `index.html` — página principal (fornecida abaixo)
- `script.js` — seu arquivo JS (já preparado para API remota)
- `README.md` — este arquivo
- `.gitignore` — recomendado

## Como configurar
1. **Preencha `CONFIG` em `script.js`**  
   - Para Google Sheets:
     ```js
     CONFIG.useApi = true;
     CONFIG.apiType = "google_sheets";
     CONFIG.sheetId = "SEU_SHEET_ID";
     CONFIG.apiKey = "SUA_API_KEY";
     CONFIG.range = "A1:Z";
     ```
   - Para endpoint JSON:
     ```js
     CONFIG.useApi = true;
     CONFIG.apiType = "json_endpoint";
     CONFIG.endpointUrl = "https://api.seuservidor.com/produtos";
     ```

2. **CORS e autenticação**  
   - Se a API exigir autenticação complexa (OAuth) ou bloquear CORS, crie um proxy/backend que faça a chamada e exponha os dados ao frontend.  
   - **Nunca** comite chaves sensíveis em repositórios públicos. Use variáveis de ambiente no backend ou GitHub Secrets.

3. **Remover upload local**  
   - Se for usar apenas API remota, remova o `<input id="inputExcel">` do `index.html` e o listener correspondente no `script.js`.

4. **Testes locais**  
   - Abra `index.html` em um servidor local (ex.: `live-server`, `http-server`) para evitar problemas de CORS ao testar fetchs.
   - No console do navegador, rode `await carregarProdutos(true)` após preencher `CONFIG` para testar.

## Formato esperado dos dados
- **Google Sheets**: primeira linha com cabeçalhos (ex.: `Código`, `Produto`, `Valor`, `Quantidade`). O script converte cabeçalhos para minúsculas e normaliza acentos.
- **Endpoint JSON**: deve retornar um **array de objetos**. Exemplo:
  ```json
  [
    { "Código": "001", "Produto": "Parafuso", "Valor": "1.50", "Quantidade": 100 },
    { "Código": "002", "Produto": "Porca", "Valor": "0.50", "Quantidade": 200 }
  ]
