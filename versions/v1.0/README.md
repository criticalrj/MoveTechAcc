# Move TA – realocação de Technical Accounts (STAR IPS 3.62.x)

Substitui a planilha `Move TA_3.53_r5.xls` (VBA + componentes COM do STAR) por:

| Arquivo | O que é |
|---|---|
| `usp_IPS_MoveTA.sql` | Cria `dbo.usp_IPS_MoveTA` (move 1 TA origem → destino) e `dbo.usp_IPS_RetireTA` (desativa uma TA vazia). Rode uma vez no banco do IPS. |
| `Move-TA.ps1` | Interface (PowerShell/WinForms). Conecta no SQL Server, mostra a árvore de TAs da planta, pergunta como definir o destino, simula e executa. Oferece instalar as procs se não existirem. |
| `Move-TA.cmd` | Atalho: abre a interface (`powershell -STA`). |

Requisitos: Windows com PowerShell 5.1, acesso ao SQL Server do IPS (Windows ou SQL login), IPS 3.62.x (testado em 3.62.14 e 3.62.18; compat level 100 ou maior).

## Uso

1. `Move-TA.cmd` → servidor, banco, autenticação → **Conectar** → escolha a **planta**.
2. Marque na árvore as TAs de origem (ex.: 801.01, 801.04, 801.10). Os subaccounts são marcados junto (opção "Incluir subaccounts").
3. **Definir destino…** – a tool pergunta:
   * **Apenas trocar o início do código**: `De 801.` `Para 701.` → 801.01 → 701.01, 801.01.01 → 701.01.01, 801.04 → 701.04 …
   * **Escolher o código de cada TA de topo**: grade com Novo código (e nome opcional) por TA; os subaccounts seguem o novo código do pai (801.01 → 701.05 faz 801.01.02 → 701.05.02).
4. A grade "Destino" mostra cada par com a **situação**:
   * `sera criada` – destino não existe; será criado em `ST_Function` + `ST_PlantFunction` + objeto, copiando as propriedades da origem.
   * `existe na classe, sera ativada na planta` – já existe no plano de contas da classe, só falta na planta.
   * `MERGE - ja existe na planta: <nome>` (amarelo) – o destino já é uma TA ativa com outro conteúdo; as ligações da origem são **somadas** às dela. Confira se é isso mesmo (no teste, 801.01 → 701.01 caía em "Fuel oil transfer pumps"; nesse caso use "Escolher o código").
   * `INVALIDO` (vermelho) – bloqueia a execução.
   A coluna **Ligações** resume o que está pendurado na TA (jobs, spares, counters, WH, where-list).
5. **Simular**: roda tudo numa transação e faz ROLLBACK; o log mostra exatamente o que seria feito (tabela, ação, linhas).
6. **EXECUTAR**: mesma transação, com COMMIT. Ou grava tudo, ou nada. **Salvar log…** exporta CSV.

Ordem interna: pares em ordem de tamanho do código (pais antes dos filhos, para o pai existir quando o filho é criado); depois uma segunda passada desativa as origens dos filhos para os pais (um pai só fica vazio depois que todos os subaccounts saíram).

## O que a procedure move (3.53 → 3.62)

| 3.53 (planilha) | 3.62.x (proc) | Observação |
|---|---|---|
| `ST_PlantFunctionObject` | `ST_PlantFunctionObject` | jobs 320, spare parts 221, counters 210, equipamentos… Soft-delete da linha antiga + insert/reativação da nova (PK nunca é alterada – seguro p/ replicação). |
| `ST_WPWhere (WP, FuncClass, Function)` | `ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID)` | `OBJECT_ID` = objeto da TA (`ST_PlantFunction.PLANTFUNC_OBJECT_ID`, `ST_Object` tipo 2). Mantém MAIN/OK/comments/status. A tabela antiga só é tocada se ainda tiver linhas. |
| `ST_WHWhere` | `ST_WHWHERE2` | idem; colunas desnormalizadas (`WHWHERE_FUNCTION_ID`, `_PLANTFUNC_ID`…) só são preenchidas se já estavam preenchidas na origem. |
| `ST_WorkHistory.WH_Function_Id/Name` | igual + `PLANTFUNC_ID/LOCATION` | só histórico ativo cuja TA principal é a origem. |
| `ST_POAITEM2.Function_Id` | igual (+ filtro por planta via `ST_Object.ObjClass_Id`) | |
| — | `ST_MATONHANDSTAT.MatOH_Def_Function_Id`, `ST_MATTRANS2.MATTR_FUNCTION_ID` | TA default do material e histórico de transações (opcional). |
| — | `ST_BudItem`, `ST_PlantCost`, `ST_UserGroupRestrictFunc` | update direto; item de orçamento que já exista no destino é reportado como `ATENCAO` e não é movido. |
| `CopyTA` (propriedades) | automático ao criar o destino | todas as colunas de `ST_PlantFunction` e `ST_PLANTFUNCEQ`. |
| — | `ST_Function` | cria/reativa o destino; pai = mesma profundidade da origem (801.01 → pai 801, logo 701.01 exige 701 existente); `Function_Length`/`Function_OnLower` recalculados. Erro claro se o grupo pai não existir. |
| — | desativar origem | `ST_PlantFunction`, `ST_PLANTFUNCEQ`, `ST_Object` (Rec_Deleted = -1) só se nada ativo restar; na classe (`ST_Function`) só se nenhuma outra planta usar e não houver subaccounts ativos. |

Tabelas que apenas **aparecem no log como `VERIFICAR`** (referência por `Function_Id` sem contexto seguro de planta): `ST_NonConformity`, `ST_GuaranteeClaim`, `ST_INVAHead`, `ST_EventConsequence`, `ST_EventLessonLearned`, `ST_WorkPlanProject`, `ST_WorkHistoryProject`, `ST_ASSEMBLYNODE`, `ST_SELECTGROUP`.

## Avisos

* **Classe compartilhada**: `ST_Function` é por classe (`FuncClass_Id`). Criar 701.04 na classe faz a TA aparecer para todas as plantas da classe (como no IPS). A movimentação de ligações é sempre por planta.
* **Replicação**: os triggers padrão do IPS (`TI_/TU_`) rodam normalmente e alimentam `ST_UpdateReplTableLog`; nada é apagado fisicamente nem tem PK alterada. Mesmo assim, em ambiente replicado (main office/satélites) faça no hub e confira a replicação. A tool **não** grava em `ST_FUNCUPDATE`/`ST_WPUPDATETA` (logs de "update from master plant" do próprio IPS).
* **Backup** antes de executar em produção; a simulação usa a mesma transação, então o que o log mostra é o que será gravado.
* Teste feito em `ADEF` (3.62.14, planta FMHUB "Flavio Test"): 801.01/801.04/801.10 + 33 subaccounts → 701.xx, 36 pares, 611 linhas de log, consistência jobs ↔ `ST_WPWHERE2` (MAIN) ok, rollback verificado.
* Execução real (COMMIT) em `ADEF_MOVETA` (restore de FMTS.bak, arquivo 3) com dados de `testdata_ADEF_MOVETA.sql`: 38 jobs, 42 linhas em `ST_WPWHERE2`, 10 históricos, 12 linhas em `ST_WHWHERE2`, incluindo where-lists cruzadas 801↔701. Depois do move: nenhuma ligação ativa em 801.01/04/10, todas as 101 ligações batem com o antes traduzido 801→701 (a única linha a menos é um merge esperado em `ST_WHWHERE2`), jobs com MAIN correto e `WH_Function_Id` coerente com o MAIN de `ST_WHWHERE2`.
* Ao rodar scripts pelo `sqlcmd`, use `-I` (QUOTED_IDENTIFIER ON); sem isso alguns INSERTs falham por causa de índices filtrados. A interface e o .NET já usam ON.

## Uso direto via SQL (sem GUI)

```sql
-- simulação de um par
EXEC dbo.usp_IPS_MoveTA @Plant_Id='FMHUB', @From_Function_Id='801.04', @To_Function_Id='701.04', @DryRun=1;

-- vários pares, tudo ou nada
BEGIN TRAN;
EXEC dbo.usp_IPS_MoveTA 'FMHUB','801.04','701.04',    NULL,1,1,1,1,1, 0,0,0,NULL, 0;
EXEC dbo.usp_IPS_MoveTA 'FMHUB','801.04.01','701.04.01',NULL,1,1,1,1,1, 0,0,0,NULL, 0;
EXEC dbo.usp_IPS_MoveTA 'FMHUB','801.04.02','701.04.02',NULL,1,1,1,1,1, 0,0,0,NULL, 0;
EXEC dbo.usp_IPS_RetireTA 'FMHUB','801.04.02',1;  -- filhos primeiro
EXEC dbo.usp_IPS_RetireTA 'FMHUB','801.04.01',1;
EXEC dbo.usp_IPS_RetireTA 'FMHUB','801.04',1;
-- revise os result sets; depois:
COMMIT;   -- ou ROLLBACK
```
