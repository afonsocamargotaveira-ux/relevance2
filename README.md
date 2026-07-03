# relevance – Avaliação de Aderência (versão Vercel)

Este projeto substitui o app Streamlit original por:

- **Frontend estático** (`public/`): HTML + CSS + JS puro, reproduzindo o mesmo
  layout, cores e textos da interface original.
- **API serverless em Python** (`api/index.py`): Flask, com toda a lógica de
  parsing de termos, pontuação (`score_relevance`), leitura/geração de xlsx e
  consulta ao Inteiro Teor portada do `app.py` original, sem alterações na
  regra de negócio.

Streamlit não roda no Vercel porque depende de um servidor com estado
persistente (`st.session_state`) e conexão contínua (WebSocket) — o Vercel só
executa funções serverless sem estado. Por isso a reescrita foi necessária.

## Estrutura

```
relevance-vercel/
├── api/
│   └── index.py        # rotas /api/columns e /api/process
├── public/
│   ├── index.html
│   ├── styles.css
│   └── app.js
├── requirements.txt
├── vercel.json
└── README.md
```

## Deploy

1. Instale a CLI da Vercel (`npm i -g vercel`) ou conecte o repositório
   diretamente pelo painel da Vercel.
2. Dentro da pasta `relevance-vercel`, rode:
   ```
   vercel
   ```
   e siga as instruções (ou faça `git push` para um repositório conectado ao
   Vercel).
3. Nenhuma variável de ambiente é necessária.

## Pontos de atenção específicos do Vercel

- **Tempo de execução**: funções serverless têm limite de duração. O
  `vercel.json` já pede `maxDuration: 60`, mas esse valor só é respeitado em
  planos pagos (no plano Hobby o limite é 10s). A etapa de consulta ao
  **Inteiro Teor** (download de PDFs da Câmara) é a mais lenta — se o plano
  não suportar 60s, desmarque a opção "Consultar Inteiro Teor" na interface
  para planilhas grandes, ou rode essa etapa localmente.
- **Tamanho do upload**: o corpo de uma requisição serverless na Vercel tem
  limite (por padrão ~4,5 MB). Planilhas muito grandes podem precisar de
  outra estratégia de upload (ex.: upload direto para um storage e
  processamento em background).
- **Sem cache entre requisições**: o `st.session_state` foi removido; o cache
  de URLs do Inteiro Teor agora vive apenas durante uma única requisição.
- **Sem impressão incremental de progresso**: como a resposta só chega ao
  final do processamento, a barra de progresso da versão Streamlit foi
  substituída por um spinner simples no botão "Avaliar".

## Testando localmente

```bash
pip install -r requirements.txt
python api/index.py          # sobe o Flask em http://localhost:5000
```

E sirva a pasta `public/` com qualquer servidor estático (ex.: `npx serve public`),
ajustando as URLs de fetch em `app.js` para `http://localhost:5000/api/...` se
não estiver usando o proxy do `vercel dev`.

Ou, de forma mais simples, use o próprio Vercel:

```bash
vercel dev
```

que já expõe `/api/*` e os arquivos estáticos juntos, exatamente como em
produção.
