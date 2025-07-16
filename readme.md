# Script Gerador de Ranking de Vendas ⚡

Este script em **Node.js** gera um ranking de vendedores baseado no **valor total de vendas instaladas** no mês selecionado, comparando os resultados com metas previamente configuradas.

---

## ⚙️ Funcionalidade

- 📁 Lê os dados a partir de um arquivo `EXPORTAR.XLS` ou nome personalizado na CLI.
- 📅 Permite selecionar o **mês** e **ano** da análise via terminal.
- 📊 Calcula:
  - Vendas instaladas
  - Valor total das vendas instaladas
  - Percentual de meta atingida
- 📝 Gera dois relatórios automaticamente:
  - `ranking_MES_ANO.txt`: resumo em texto para compartilhamento.
  - `ranking_MES_ANO.html`: versão visual estilizada para apresentação.

---

## ▶️ Como usar

1. Clone ou baixe este repositório.
2. Instale as dependências:
   ```bash
   npm install
   ``` 

3. Coloque o arquivo EXPORTAR.XLS na mesma pasta.

4. Execute o script:
    ```bash
   node ranking.js
   ``` 

## 🧠 Comandos úteis no terminal

`gerarRankingVendas()` – Executa a análise completa.

`verDetalhesVendedor('Nome')` – Mostra detalhes individuais de um vendedor.

`verResumoMetas()` – Mostra um resumo geral das metas atingidas.

## ✏️ Configurações
Os nomes dos vendedores e metas estão no objeto CONFIG, no topo do arquivo ranking.js. Edite conforme necessário.

## 📦 Dependências
- xlsx
- fs
- readline-sync

```bash
npm install xlsx readline-sync
   ``` 
   ___

   Feito com 💙 para análise interna de desempenho comercial.