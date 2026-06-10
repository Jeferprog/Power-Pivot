# Power-Pivot

Dashboard de liberações com mapeamento de agências/gerentes por conta.

## Atualização automática da Base de Agências

A base de origem é um arquivo CSV no Google Drive (separado por `;`) com o cabeçalho:
`nome_completo; conta; cpf_cnpj; data_nascimento; m_carteira_responsavel; nm_carteira; nm_agencia`.

O fluxo recomendado usa um gatilho diário no Google Apps Script que copia **somente os
campos necessários** para uma planilha Google intermediária ("planilha base"):

| Coluna no CSV de origem    | Coluna na planilha base |
| -------------------------- | ----------------------- |
| `conta`                    | Conta                   |
| `nm_agencia`               | Nome Agência            |
| `m_carteira_responsavel`   | Gerente de Contas       |

O dashboard sincroniza a partir dessa planilha — os dados sensíveis
(CPF/CNPJ, data de nascimento) não chegam ao navegador.

```
CSV no Drive  ──(gatilho diário do Apps Script)──►  Planilha base  ──(sincronização ao abrir / a cada 5 min)──►  Dashboard
```

### Passo a passo

1. Acesse <https://script.google.com>, crie um projeto e cole o conteúdo de
   [`google-apps-script/atualizar_base.gs`](google-apps-script/atualizar_base.gs).
2. Execute a função `atualizarBase` uma vez (autorize as permissões). O log mostra a
   URL da planilha base criada automaticamente.
3. Execute a função `instalarGatilhoDiario` — a base passa a ser atualizada todos os
   dias (horário configurável em `CONFIG.HORA_GATILHO`).
4. Compartilhe a planilha base como **"Qualquer pessoa com o link" (Leitor)**.
5. No dashboard, no card **Base de Agências → Atualização Automática (Google Drive)**,
   cole o link da planilha base e clique em **Conectar**.

O dashboard então sincroniza a base automaticamente ao abrir a página e a cada
5 minutos, mantendo o botão **Atualizar Base** (arquivo `.xlsx`/`.xls`/`.csv`)
como alternativa manual.

> Alternativa sem Apps Script: o dashboard também aceita o link/ID do próprio arquivo
> CSV no Drive (compartilhado publicamente) — nesse caso ele baixa o arquivo completo
> diretamente, sem o filtro de campos.
