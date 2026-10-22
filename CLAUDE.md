# MoveTechAcc – contexto para o Claude

Ferramenta para **realocar, criar e renomear Technical Accounts (TA)** no STAR IPS 3.62.x (EAM marítimo da Star Information Systems). Substitui a planilha VBA "Move TA_3.53_r5.xls", que não cobria as tabelas novas da 3.62. Autor: Flávio Miguel. Versão atual: **2.3**.

## Arquivos
- `Move-TA.ps1` – app WPF em PowerShell 5.1 (~2.600 linhas), bilíngue PT/EN. `$script:Version = '2.3'`.
- `usp_IPS_MoveTA.sql` – procs `usp_IPS_MoveTA`, `usp_IPS_RetireTA`, `usp_IPS_CreateTA`, `usp_IPS_RenameTA`. Todas participam da transação do chamador.
- `testdata_ADEF_MOVETA.sql` – árvore de teste 7/701 e 8/801 com jobs, WH, WPWHERE2, WHWHERE2.
- `build/Build-Exe.ps1` + `build/MoveTA.Launcher.cs` – geram `dist/MoveTA.exe` (exe portátil com ps1 e sql embutidos, compilado com o csc.exe do .NET 4.x).
- `Gerar-Exe.cmd` – dois cliques no Windows rodam `build/Build-Exe.ps1` e geram `dist/MoveTA.exe`.
- **Não compilar o exe com Mono na nuvem:** o exe gerado com `mcs` (testado na 2.3) não abre no PC do usuário, sem nenhum erro. Só o build com o csc.exe do Windows funciona.
- `versions/` – v1.0 (WinForms) e v2.0. `docs/screenshots/` – prints.
- `README.md` – manual completo em PT, com histórico das versões.

## Modelo de dados do IPS 3.62 (essencial)
- A TA é um objeto: `ST_PlantFunction.PLANTFUNC_OBJECT_ID` → `ST_Object` (ObjType 2, Obj_Code = Function_Id). Objeto novo via `st_CreateObjectId`.
- `ST_WPWHERE2` / `ST_WHWHERE2` usam (WP/WH_OBJECT_ID, OBJECT_ID = objeto da TA). **Sempre cobrir essas duas tabelas** (exigência do usuário; a planilha antiga não cobria). `ST_WPWhere`/`ST_WHWhere` legadas estão vazias; as views `VI_WPWHERE`/`VI_WHWHERE` leem das tabelas com 2.
- `ST_PlantFunctionObject` liga jobs (320), spares (221) e counters (210).
- `ST_Function` é por classe (FuncClass_Id): hierarquia por Function_Father; Function_Length = LEN(código); Function_OnLower = nº de filhos.
- Pai do destino: o app sempre passa `@To_Father` (último parâmetro de `usp_IPS_MoveTA`, desde 2.3). Código livre: só precisa começar com o pai e ser mais longo, e não pode parecer subaccount de outra TA.
- Soft delete com `Rec_Deleted = -1`. **Nunca alterar PK nem apagar fisicamente** (replicação): desativa a linha antiga e insere/reativa a nova.
- Pai de um destino novo = mesma profundidade da origem (801.01 → pai 801, logo 701.01 exige 701).

## Lógica do app
- Duas árvores (estrutura atual / nova), arrastar e soltar, tudo vira item de um **plano** (move / create / rename).
- Execução numa única transação SqlClient: criar e mover em ordem crescente de tamanho do código, depois renomear, depois desativar origens em ordem decrescente. **Simular = ROLLBACK**; Executar = COMMIT, só depois de simulação sem erro. Log CSV em `logs/`.
- Settings em `%APPDATA%\MoveTA\settings.json` (lang, server, database).
- Textos via `$script:Strings` (pt/en) + DynamicResource para trocar idioma ao vivo.

## Regras do código (exigência do usuário)
- **Comentários em inglês** no app, no launcher, no build e nas procedures. Textos de interface continuam bilíngues; o log das procs continua em português.
- **Toda correção nas procedures entra no cabeçalho de `usp_IPS_MoveTA.sql`** (CHANGE LOG com data, versão e o fix).
- Ao mudar qualquer procedure, subir a marca `-- MoveTA-Procs-Version: X.Y` dentro de **cada uma das 4 procs**. O app compara essa marca com o arquivo embutido e atualiza as procs sozinho (ao conectar e antes de simular/executar).
- O app pode remover as procs do banco do cliente: botão na tela de conexão ou opção "ao fechar" (`dropProcsOnClose` no settings.json).

- **Sempre entregar o app para gerar o executável** ao fim de cada mudança: zip `MoveTA_<versão>.zip` com a pasta `MoveTA_<versão>/` contendo `Move-TA.ps1`, `usp_IPS_MoveTA.sql`, `Move-TA.cmd`, `Gerar-Exe.cmd`, `README.md`, `build/Build-Exe.ps1` e `build/MoveTA.Launcher.cs`, enviado como anexo. O usuário roda `Gerar-Exe.cmd` e usa o `dist\MoveTA.exe` gerado. O usuário usa o exe, não o script.

## Armadilhas (importante)
- `Move-TA.ps1` precisa ser salvo em **UTF-8 com BOM** (PowerShell 5.1).
- `usp_IPS_MoveTA.sql` deve ser **só ASCII** (sem acentos): sqlcmd e o app leem sem BOM.
- PowerShell: `$(if ...)` desmonta coleções – usar if/else normal; scriptblocks de eventos WPF só enxergam variáveis `$script:`.
- sqlcmd precisa de `-I` (QUOTED_IDENTIFIER, índices filtrados) e `-C` (certificado).
- Ao mudar o app: atualizar `$script:Version`, a versão no `MoveTA.Launcher.cs` (const + atributos) e o README; depois gerar o exe no Windows com `Gerar-Exe.cmd`.

## Limites da sessão na nuvem
- A nuvem é Linux: **não roda WPF, PowerShell 5.1 nem SQL Server do IPS**. Dá para editar código, SQL e README, revisar e abrir PRs. **O exe não sai daqui**: é gerado no Windows com `Gerar-Exe.cmd`.
- Testes de verdade e o build do `MoveTA.exe` são feitos no PC do Flávio (localhost\SQL2025, banco sandbox `ADEF_MOVETA`, planta FMHUB) com `Move-TA.ps1 -UiTest` e simulação.
- Para validar sintaxe na nuvem dá para usar `pwsh` (PowerShell 7), sabendo que WPF não carrega.

## Pendências conhecidas
- Colocar um print (`docs/screenshots/MoveTA2_tela_pt.png`) no topo do README.
