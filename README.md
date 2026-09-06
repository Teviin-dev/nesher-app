# NESHER — Liberação de Chapas, ligado à planilha

Esse pacote conecta a tela de Liberação de Chapas direto à sua planilha
`Controle_Chapas_MDF_FUNDOS.xlsx`, na aba **CHAPAS POR LOTE** (tabela `Tabela10`).

- A tela **lê** a planilha e mostra as liberações recentes + os totais (chapas
  liberadas hoje, OPs atendidas, total histórico).
- O formulário **grava** uma nova linha direto na planilha (OP, Tipo de Chapa,
  Quantidade e Perda %). Fábrica e M² da chapa são calculados automaticamente,
  do mesmo jeito que a planilha já faz hoje.
- A tela se atualiza sozinha a cada 15 segundos — então se alguém editar a
  planilha diretamente (Excel), a mudança aparece na tela também.

## Como rodar

1. Instale as dependências (Python 3.10+):
   ```
   cd backend
   pip install -r requirements.txt
   ```

2. Coloque a planilha real no lugar certo. Por padrão o backend procura o
   arquivo `Controle_Chapas_MDF_FUNDOS.xlsx` na pasta `backend/` (já incluí
   uma cópia dela aqui). Se a planilha oficial de vocês estiver em outro
   caminho — por exemplo, no servidor interno da empresa — aponte pra ela
   com uma variável de ambiente antes de iniciar:

   Windows (PowerShell):
   ```
   $env:PLANILHA_PATH="\\SERVIDOR\PPCP\Controle_Chapas_MDF_FUNDOS.xlsx"
   python main.py
   ```

   Linux/Mac:
   ```
   PLANILHA_PATH="/caminho/da/rede/Controle_Chapas_MDF_FUNDOS.xlsx" python main.py
   ```

3. Rode o servidor:
   ```
   python main.py
   ```
   Ele sobe em `http://localhost:8000` — abra esse endereço no navegador e a
   própria tela já aparece (o backend serve o HTML e a API juntos).

   Se quiser que outras pessoas da empresa acessem pela rede interna, use o
   IP da máquina onde ele está rodando, ex: `http://192.168.0.50:8000`.

## Pontos de atenção

- **Um arquivo, um "dono" por vez.** A planilha só pode ser editada por uma
  aplicação de cada vez sem risco de conflito. Se alguém estiver com o Excel
  aberto no arquivo enquanto o backend tenta gravar, o Windows pode bloquear
  a escrita — nesse caso o formulário mostra erro e nada é gravado (a
  planilha não fica corrompida).
- **Cálculos automáticos:** Fábrica é derivada da primeira letra da OP (N,
  I/M, L/U, A/E, V) e M² da chapa vem do "Tipo de Chapa" escolhido — os
  mesmos mapeamentos que já existem nas abas REFERÊNCIA/LISTA da planilha.
  Se vocês cadastrarem um tipo de chapa novo na planilha, adicione-o também
  no dicionário `M2_POR_TIPO_CHAPA` em `backend/main.py`.
- **ERP:** os campos ligados ao ERP (coluna I em diante) ficam em branco nas
  liberações feitas pela tela — a planilha já lida com isso hoje (muitas
  linhas também estão em branco ali).
- Isso está rodando localmente, como protótipo. Pra deixar isso no ar de
  forma permanente no servidor da empresa (ligado, reiniciando sozinho se
  Deploy Vercel conectado
  cair, etc.), o ideal é rodar como serviço do Windows/Linux — posso te
  ajudar a montar isso quando vocês estiverem prontos para produção.
