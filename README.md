# Move TA – realocação de Technical Accounts (STAR IPS 3.62.x)

## MoveTA.exe (portátil, pendrive)

`dist\MoveTA.exe` é um arquivo único: leva o app e as procedures embutidos. Copie só ele para o pendrive.

* Ao abrir, extrai os arquivos para `%TEMP%\MoveTA_<versão>` e roda o app com o PowerShell 5.1 do Windows, sem console.
* Os logs (CSV) ficam na pasta `logs` ao lado do .exe, no próprio pendrive.
* Idioma, servidor e banco da última conexão ficam em `%APPDATA%\MoveTA\settings.json` de cada máquina.
* Requisitos: Windows 10/11 com PowerShell 5.1 e .NET Framework 4.x (padrão do Windows). Não instala nada.
* Se não abrir: rode `MoveTA.exe /debug` para ver o erro. A causa mais comum é política (GPO) que bloqueia scripts PowerShell.
* Parâmetros passam direto para o app, por exemplo `MoveTA.exe -Lang en -Server SRV\SQL -Database IPS`.
* Para gerar de novo depois de mudar o app: `powershell -NoProfile -ExecutionPolicy Bypass -STA -File build\Build-Exe.ps1`. No Linux, com o Mono: `build/build-exe-linux.sh`.

## Versão 2.3 (atual)

* **Código livre ao mover ou mudar código:** o código novo só precisa começar com o grupo pai e ser mais longo. Não precisa mais ter o mesmo tamanho de nível da origem. Exemplo: 731 solto em 9999 pode virar 9999.731, e os filhos seguem (731.100 → 9999.731.100). Antes o diálogo exigia o formato 9999N.
* A tool recusa um código que pareça subaccount de outra TA, por exemplo 99991.731 dentro de 9999 quando 99991 existe.
* **Procedures atualizadas sozinhas:** cada procedure leva a marca `MoveTA-Procs-Version`. Ao conectar e antes de simular ou executar, a tool compara a marca do banco com a do `usp_IPS_MoveTA.sql` embutido. Se faltar alguma procedure, ela pergunta antes de instalar. Se estiverem desatualizadas, ela atualiza e avisa na tela de conexão.
* **Remover as procedures do banco do cliente:** botão **Remover procedures** na tela de conexão, ou a opção **Remover as procedures do banco ao fechar a tool**, que fica salva em `settings.json`. Nenhum dado do IPS é alterado. Se a tool for usada de novo, as procedures voltam a ser instaladas.
* `usp_IPS_MoveTA` ganhou o parâmetro opcional `@To_Father`, o último da lista. A tool sempre informa o pai. Sem ele, vale a regra antiga de mesma profundidade da origem.
* Comentários do código do app, do launcher e das procedures passaram para inglês. O cabeçalho de `usp_IPS_MoveTA.sql` tem o histórico de correções com data.
* Em ambiente com mais de um usuário da tool no mesmo banco, não use a remoção ao fechar: ela apaga as procedures que outra pessoa pode estar usando.

## Versão 2.2

* **Renomear ou mudar código:** o mesmo diálogo agora muda o código também, e os filhos seguem. Em TA criada no plano, os itens dentro dela são reescritos. No topo de uma estrutura movida, muda o destino. Em TA já salva, vira renumeração dentro do mesmo pai. Filhos de estrutura movida têm o código travado: edite o topo.
* **Sobre:** botão no trilho com versão, autoria (Flávio Miguel) e ambiente (conexão, versão do IPS, procedures, pasta de logs).
* O MoveTA.exe leva versão 2.2 e autoria nas propriedades do arquivo.

## Versão 2.1

Tudo vira item de um **plano** antes de ir para o banco:

* **Mover:** arrastar e soltar, inclusive sobre uma TA que ainda não foi salva (criada ou movida no próprio plano).
* **Nova TA:** botão direito numa TA → "Nova TA aqui dentro…", ou o botão **Nova TA** na barra. A tool sugere o próximo código livre no formato dos irmãos.
* **Renomear:** botão direito → "Renomear…", o botão **Renomear** ou F2. O nome mora na classe (`ST_Function`), então muda em todas as plantas da classe, e o diálogo avisa quantas. Em TA ainda não salva, só troca o nome com que ela será criada.
* **Remover do plano:** um item que é destino de outros leva os dependentes junto, com aviso.
* **Executar:** uma transação, na ordem criar/mover (pais antes dos filhos), depois renomear, depois desativar as origens vazias.
* **Confirmações:** todas no visual da tool, sem caixas padrão do Windows.

Novas procedures em `usp_IPS_MoveTA.sql`: `usp_IPS_CreateTA` e `usp_IPS_RenameTA`. A tool oferece instalar ao conectar. A versão 2.0 ficou em `versions/Move-TA_v2.0_bilingue.ps1`.

## Versão 2.0

Interface WPF no visual do FM Terminals (Fluent): barra azul com busca, trilho Realocar / Histórico / Conexão.

1. Abre já conectado (padrão `localhost\SQL2025` / `ADEF_MOVETA`; troque em Conexão).
2. Arraste uma TA da **Estrutura atual** e solte sobre o **novo grupo pai** na **Nova estrutura**. A TA leva tudo que está abaixo dela.
3. O diálogo pergunta como aplicar os códigos: **trocar só o início** (801.10 → 701.10), **próximo código livre** (→ 701.11) ou **escolher o código**. A prévia mostra cada TA com a situação: NOVA, REATIVAR, MERGE ou conflito.
4. O movimento entra no **Plano**; empilhe quantos quiser. A nova estrutura mostra as TAs novas em azul e as de origem riscadas.
5. **Simular** roda tudo numa transação e desfaz. **Executar** só habilita depois de uma simulação sem erro.
6. Cada simulação e execução vira um CSV em `logs/`, listado na tela Histórico.

Idioma: o botão PT/EN na barra de título troca português ↔ inglês na hora, sem perder o plano, e a escolha fica salva em `%APPDATA%\MoveTA\settings.json`. Também dá para abrir com `-Lang en`. O detalhe do log vem das procedures do banco e continua em português.

Atalhos: Enter na busca localiza a TA; Esc limpa; Ctrl+Z desfaz o último movimento; F5 recarrega. A versão 1.0 (WinForms) está em `versions/v1.0` e em `versions/MoveTA_v1.0.zip`.

---

## Versão 1.0

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
| — | `ST_Function` | cria/reativa o destino; pai = `@To_Father` quando informado (a tool sempre informa), senão mesma profundidade da origem (801.01 → pai 801, logo 701.01 exige 701 existente); `Function_Length`/`Function_OnLower` recalculados. Erro claro se o grupo pai não existir. |
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

-- código livre (2.3): pai informado no último parâmetro, @To_Father
EXEC dbo.usp_IPS_MoveTA @Plant_Id='FMHUB', @From_Function_Id='731', @To_Function_Id='9999.731', @DryRun=1, @To_Father='9999';
```
