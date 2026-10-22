<#
.SYNOPSIS
    Move TA v2.2 - realocação, criação, renomeação e renumeração de Technical Accounts no STAR IPS 3.62.x.
    Move TA v2.2 - relocate, create, rename and renumber Technical Accounts in STAR IPS 3.62.x.

.DESCRIPTION
    Visual no padrão do FM Terminals (Fluent / novo Outlook). Português e inglês (botão PT/EN).

    Tudo entra num PLANO antes de ir para o banco:
      - Mover: arraste uma TA (com a estrutura abaixo dela) e solte sobre o novo pai, inclusive
        sobre uma TA que ainda não foi salva (criada ou movida no próprio plano).
      - Nova TA: botão direito numa TA (ou botão "Nova TA") cria uma TA dentro dela.
      - Renomear: botão direito, botão "Renomear" ou F2.
    Simular roda tudo numa transação com ROLLBACK; Executar grava tudo (COMMIT), na ordem:
    pais antes dos filhos (criar/mover), depois renomear, depois desativar as origens vazias.

    Procedures (usp_IPS_MoveTA.sql, instaladas pela tool se faltarem):
    usp_IPS_MoveTA, usp_IPS_RetireTA, usp_IPS_CreateTA, usp_IPS_RenameTA.
    Versões anteriores: versions\v1.0 (WinForms) e versions\Move-TA_v2.0_bilingue.ps1.
#>
[CmdletBinding()]
param(
    [string]$Server   = 'localhost\SQL2025',
    [string]$Database = 'ADEF_MOVETA',
    [string]$Plant    = '',
    [ValidateSet('', 'pt', 'en')][string]$Lang = '',
    [string]$LogDir = '',
    [switch]$NoAutoConnect,
    [switch]$UiTest,
    [string]$ScreenshotDir,
    [string]$DemoCreate       = '7.02',
    [string]$DemoCreateParent = '7',
    [string]$DemoCreateName   = 'Grupo teste 7.02',
    [string]$DemoSrc          = '801.11.02',
    [string]$DemoTarget       = '7.02',
    [string]$DemoRename       = '801.11',
    [string]$DemoRenameTo     = 'Fuel oil transfer pumps (renomeada)'
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Data

$script:Version      = '2.2'
$script:ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ProcFile     = Join-Path $script:ScriptDir 'usp_IPS_MoveTA.sql'
$script:LogDir       = $(if ($LogDir) { $LogDir } else { Join-Path $script:ScriptDir 'logs' })
$script:SettingsFile = Join-Path $env:APPDATA 'MoveTA\settings.json'

# ============================================================================ idiomas / languages
$script:Strings = @{
pt = @{
    search_hint = 'Pesquisar TA por código ou nome e tecle Enter'
    rail_move = 'Realocar'; rail_hist = 'Histórico'; rail_conn = 'Conexão'
    tt_min = 'Minimizar'; tt_max = 'Maximizar'; tt_close = 'Fechar'; tt_lang = 'Switch to English'
    move_title = 'Realocar Technical Accounts'
    move_sub = 'Arraste uma TA (ela leva tudo que está abaixo) e solte sobre o novo pai. Botão direito numa TA cria uma TA nova dentro dela ou renomeia.'
    plant = 'Planta'
    btn_sim = 'Simular'; tt_sim = 'Roda o plano numa transação e desfaz no fim'
    btn_exec = 'Executar'; tt_exec = 'Grava o plano (habilita depois de uma simulação sem erro)'
    btn_newta = 'Nova TA'; tt_newta = 'Cria uma TA dentro da TA selecionada (entra no plano)'
    btn_rename = 'Renomear'; tt_rename = 'Renomeia a TA selecionada (entra no plano) · F2'
    btn_undo = 'Desfazer'; tt_undo = 'Desfaz o último item do plano (Ctrl+Z)'
    btn_clear = 'Limpar plano'; btn_reload = 'Recarregar'; tt_reload = 'Recarrega a planta do banco (F5)'; btn_opts = 'Opções'
    mn_new = 'Nova TA aqui dentro…'; mn_rename = 'Renomear ou mudar código…'
    opt_h1 = 'O que vai junto com a TA'
    opt_always = 'Sempre: jobs, spares, counters (PFO) e where-list dos jobs (ST_WPWHERE2)'
    opt_wh = 'Histórico (ST_WorkHistory e ST_WHWHERE2)'; opt_po = 'Linhas de MR/PO (ST_POAITEM2)'
    opt_mat = 'Materiais: TA default e transações'; opt_other = 'Orçamento, custos e restrições de grupo'
    opt_h2 = 'Depois de mover'; opt_retire = 'Desativar as TAs de origem que ficarem vazias'
    opt_retclass = 'Desativar também na classe, se nenhuma outra planta usar'; opt_items = 'Listar cada objeto movido na atividade'
    src_title = 'Estrutura atual'; src_sub = 'Arraste daqui'; src_count = '{0} TAs'
    dst_title = 'Nova estrutura'; dst_hint = 'Solte sobre o pai, mesmo que ele ainda não esteja salvo'
    chip_new = 'NOVA'; chip_react = 'REATIVAR'; chip_merge = 'MERGE'; chip_rename = 'RENOMEAR'
    plan_title = 'Plano'; plan_sub = 'Executado na ordem, tudo ou nada'; tt_remove = 'Remover do plano'
    plan_empty_t = 'Nada no plano ainda'
    plan_empty_s = 'Arraste uma TA da estrutura atual e solte sobre o novo pai, ou use o botão direito para criar ou renomear uma TA.'
    act_title = 'Atividade'; act_none = 'Nenhuma simulação ainda'; btn_export = 'Exportar log'
    col_move = 'Item'; col_table = 'Tabela'; col_action = 'Ação'; col_rows = 'Linhas'; col_detail = 'Detalhe'
    hist_title = 'Histórico'; hist_sub = 'Cada simulação e execução fica salva em CSV na pasta logs, ao lado da tool.'
    btn_open = 'Abrir'; btn_openfolder = 'Abrir pasta'; btn_refresh = 'Atualizar'
    col_date = 'Data'; col_kind = 'Tipo'; col_plant = 'Planta'; col_file = 'Arquivo'
    hist_empty = 'Nenhum log ainda. Rode uma simulação na tela Realocar.'
    conn_title = 'Conexão'; conn_sub = 'Banco do STAR IPS 3.62 onde as TAs serão realocadas.'
    lbl_server = 'Servidor'; lbl_db = 'Banco'; lbl_auth = 'Autenticação'; auth_win = 'Windows'; auth_sql = 'SQL login'
    lbl_user = 'Usuário'; lbl_pass = 'Senha'; btn_connect = 'Conectar'; btn_install = 'Reinstalar procedures'
    dlg_title = 'Mover estrutura'; dlg_from = 'DE'; dlg_to = 'PARA DENTRO DE'; dlg_how = 'Como aplicar os novos códigos?'
    rule_prefix = 'Trocar só o início do código'; rule_next = 'Usar o próximo código livre'; rule_custom = 'Escolher o código'
    btn_cancel = 'Cancelar'; btn_add = 'Adicionar ao plano'; btn_apply = 'Aplicar'
    col_cur = 'Código atual'; col_newc = 'Novo código'; col_name = 'Nome'; col_links = 'Ligações'; col_status = 'Situação'
    dlg_sub = 'Planta {0} · a estrutura inteira vai junto ({1} TAs)'
    desc_next = '{0} → {1}, o primeiro número livre dentro de {2}'
    desc_next_na = 'Não dá para calcular um próximo número para esse formato de código.'
    desc_custom = 'formato {0} · os subaccounts seguem'
    err_prefix_same = '{0} já é o grupo pai atual: trocar o início não muda nada. Use o próximo código livre ou escolha o código.'
    err_next = 'Não foi possível calcular o próximo código livre.'; err_empty = 'Digite o novo código.'
    err_start = 'O código precisa começar com {0}, o grupo onde você soltou.'
    err_format = 'Use o formato {0}: o mesmo tamanho de nível que {1}.'; err_sep = "Depois de {0} vem '{1}'."
    err_equal = 'O código novo é igual ao atual.'; err_inside = 'A TA não pode ficar dentro dela mesma.'
    err_conflict = '{0} TA(s) em conflito: o destino já está saindo no plano ou fica fora do padrão de código.'
    prev = 'Prévia'; prev_hdr = 'Prévia · {0} TAs'
    cnt_new = '{0} nova(s)'; cnt_react = '{0} reativar'; cnt_merge = '{0} merge'; cnt_bad = '{0} conflito(s)'
    foot = 'Nada é gravado agora: o movimento entra no plano.'
    foot_merge = 'Atenção: {0} TA(s) já existem no destino e vão receber as ligações (merge).'
    rule_t_prefix = 'Trocar início'; rule_t_next = 'Próximo código livre'; rule_t_custom = 'Código escolhido'
    rule_t_create = 'Nova TA'; rule_t_rename = 'Renomear'
    card_create = 'Nova TA dentro de {0}'; card_rename = 'antes: {0}'
    info_added_t = 'Adicionado ao plano'; info_added = '{0} → {1} com {2} TAs. Arraste mais estruturas ou clique em Simular.'
    st_out = 'FORA DO PADRÃO'; st_same = 'SEM MUDANÇA'; st_conflict = 'CONFLITO'; st_merge_plan = 'MERGE (plano)'
    st_merge = 'MERGE'; st_activate = 'ATIVAR'; st_react = 'REATIVAR'; st_new = 'NOVA'; st_rename = 'RENOMEAR'
    drag_root = 'A raiz do plano de contas não pode ser movida.'
    drag_inplan = '{0} já está no plano. Remova o item para mudar o destino.'
    drag_child = '{0}, que fica dentro de {1}, já está no plano. Remova-o primeiro ou mova só o resto.'
    drag_dep = 'Há itens no plano que usam TAs dentro de {0} como destino ou que a renomeiam. Remova-os antes de mover {0}.'
    drag_cant_t = 'Não dá para arrastar essa TA'
    drop_none = 'Solte sobre o grupo pai de destino.'
    drop_leaving = '{0} está saindo no plano.'; drop_self = 'Uma TA não pode ir para dentro dela mesma.'
    drag_from = 'Arraste uma TA da estrutura atual.'
    dragging = 'Arrastando {0} com {1} TA(s): solte sobre o novo pai'; drop_here = 'Soltar {0} dentro de {1} · {2}'
    plan_count = '{0} item(ns)'; plan_sum = '{0} item(ns) no plano · {1} TAs'; sum_merge = ' · {0} merge'; sum_bad = ' · {0} conflito(s)'
    plan_emptysum = 'Plano vazio'; warn_conf = '{0} conflito(s)'; warn_merge = '{0} merge'
    plan_cascade_t = 'Removido do plano'; plan_cascade = 'Também saíram {0} item(ns) que dependiam dele (destinos que só existiam no plano).'
    search_t = 'Pesquisa'; search_found = "{0} TA(s) encontradas para '{1}'. Esc limpa."; search_none = "Nada encontrado para '{0}'."
    act_line = '{0} · {1} · {2} linhas de log'; act_sim = 'Simulação'; act_exec = 'Execução'
    resumo = '{0} TAs · {1} ligações PFO · {2} em ST_WPWHERE2 · {3} em ST_WHWHERE2 · {4} históricos · {5} alertas'
    fail_exec = 'Execução falhou, nada foi gravado'; fail_sim = 'Simulação falhou'
    sim_ok_t = 'Simulação OK, nada foi gravado'; sim_ok = '{0}. Revise a atividade e clique em Executar.'
    exec_ok_t = 'Plano gravado'; exec_ok = '{0}. Log salvo em {1}.'
    m_tas = 'TAs no plano'; m_created = 'TAs criadas / reativadas'; m_pfo = 'Jobs, spares, counters (PFO)'
    m_wp = 'Where-list jobs (ST_WPWHERE2)'; m_wh = 'Where-list histórico (ST_WHWHERE2)'; m_hist = 'Histórico (ST_WorkHistory)'
    m_renamed = 'TAs renomeadas'; m_retired = 'TAs de origem desativadas'; m_alerts = 'Alertas'
    log_commit = 'Alterações gravadas.'; log_rollback = 'Simulação: nada foi gravado.'; log_err = '{0} (transação desfeita, nada gravado)'
    log_retire = '{0} (desativar)'; log_create = '+ {0}'; log_rename = '{0} (renomear)'
    proc_missing = "As procedures da tool não existem (ou estão desatualizadas) em {0}.`nInstalar agora a partir de usp_IPS_MoveTA.sql?"
    proc_notinst = 'Procedures não instaladas.'; conn_ok = 'Conectado em {0} / {1} (IPS {2}). Procedures prontas.'
    conn_fail = 'Falha: {0}'; disconnected = 'Desconectado'; not_connected_t = 'Não conectou'
    plant_loaded_t = 'Planta {0} carregada'; plant_loaded = '{0} TAs na classe {1}. Arraste uma TA da esquerda e solte sobre o novo pai à direita.'
    plant_fail = 'Falha ao carregar a planta'; connect_first = 'Conecte primeiro.'; reinstalled = 'Procedures reinstaladas em {0}.'
    switch_plant = 'Trocar de planta descarta o plano atual.'
    close_confirm = 'Há itens no plano que não foram executados.'
    log_exported = 'Log exportado'; kind_exec = 'EXECUÇÃO'; kind_sim = 'SIMULAÇÃO'
    lang_t = 'Idioma'; lang_set = 'Interface em português.'
    cf_exec_title = 'Gravar o plano?'; cf_exec_sub = 'Tudo roda numa única transação: ou grava tudo, ou nada.'
    cf_tas = 'TAs'; cf_moves = 'Itens do plano'; cf_plant = 'Planta'; cf_db = 'Banco'
    cf_backup = 'Se este banco for produção, faça backup antes. A última simulação deste plano passou sem erro.'
    cf_go = 'Executar agora'; cf_more = '+ {0} item(ns)'
    cf_close_t = 'Fechar o Move TA?'; cf_close_btn = 'Fechar mesmo assim'
    cf_plant_t = 'Trocar de planta?'; cf_discard = 'Descartar e trocar'
    cf_proc_t = 'Instalar as procedures?'; cf_install = 'Instalar'
    ed_new_t = 'Nova Technical Account'; ed_ren_t = 'Renomear ou mudar código'
    ed_sub = 'Planta {0} · entra no plano e só grava quando você Executar'
    ed_inside = 'DENTRO DE'; ed_ta = 'TA'; ed_code = 'Código'; ed_name = 'Nome'
    ed_code_hint = 'Sugestão: próximo código livre em {0}. Pode usar outro, desde que comece com {0}.'
    ed_err_code = 'Digite o código.'; ed_err_start = 'O código precisa começar com {0} e ser mais longo.'
    ed_err_exists = '{0} já existe nesta planta.'; ed_err_planned = '{0} já está no plano.'
    ed_err_name = 'Digite o nome.'; ed_err_same = 'O nome é igual ao atual.'; ed_err_len = 'Máximo de 100 caracteres.'
    ed_st_new = 'Será criada na classe {0} e na planta {1}.'
    ed_st_act = 'Já existe na classe {0}: será só ativada nesta planta e mantém o nome da classe.'
    ed_st_react = 'Existia desativada: será reativada com este nome.'
    ed_ren_info = 'O nome fica na classe {0}: muda também em {1} outra(s) planta(s). Histórico já arquivado mantém o nome antigo.'
    ed_ren_info0 = 'O nome fica na classe {0} e no objeto da TA. Histórico já arquivado mantém o nome antigo.'
    ed_ren_ghost = 'Essa TA ainda não foi salva: as mudanças valem para quando ela for criada.'
    ed_foot = 'Nada é gravado agora: entra no plano.'
    ed_cant_moved = '{0} está saindo no plano. Mude o destino dela, não ela.'
    sel_none_t = 'Selecione uma TA'; sel_none = 'Clique numa TA da árvore (ou use o botão direito nela).'
    info_created_t = 'Nova TA no plano'; info_created = '{0} · {1} entra dentro de {2}. Você já pode soltar outras TAs nela.'
    info_renamed_t = 'Renomeação no plano'; info_renamed = '{0}: "{1}" → "{2}".'
    ed_code_lock_child = 'O código segue o topo da estrutura movida ({0}). Para mudar, edite {0}.'
    ed_code_hint_ren = 'Pode mudar o código também: os subaccounts seguem. Formato {0}.'
    ed_err_same_all = 'Nada mudou.'; ed_err_fmt = 'Use o formato {0}.'
    ed_st_recode = 'Novo código: {0} → {1} · {2} TA(s) mudam junto.'
    st_recode = 'NOVO CÓDIGO'; rule_t_renumber = 'Renumerar'
    info_recoded_t = 'Código alterado no plano'; info_recoded = '{0} → {1}: {2} TA(s) seguem o novo código.'
    rail_about = 'Sobre'; ab_ver = 'Versão {0}'; ab_by = 'CRIADO POR'; ab_env = 'Ambiente'; ab_close = 'Fechar'
    ab_desc = 'Realoca, cria, renomeia e renumera Technical Accounts no STAR IPS 3.62, levando junto jobs, spares, counters, histórico e where-lists (ST_WPWHERE2 e ST_WHWHERE2). Tudo passa por um plano, com simulação antes de gravar.'
    ab_conn = 'Conexão'; ab_ips = 'Versão do IPS'; ab_procs = 'Procedures'; ab_procs_v = '{0} de 4 instaladas'; ab_logs = 'Pasta de logs'; ab_ps = 'PowerShell / .NET'; ab_none = 'não conectado'
}
en = @{
    search_hint = 'Search a TA by code or name and press Enter'
    rail_move = 'Relocate'; rail_hist = 'History'; rail_conn = 'Connection'
    tt_min = 'Minimize'; tt_max = 'Maximize'; tt_close = 'Close'; tt_lang = 'Mudar para português'
    move_title = 'Relocate Technical Accounts'
    move_sub = 'Drag a TA (everything below it comes along) and drop it on the new parent. Right-click a TA to create a new TA inside it or to rename it.'
    plant = 'Plant'
    btn_sim = 'Simulate'; tt_sim = 'Runs the plan in a transaction and rolls it back at the end'
    btn_exec = 'Execute'; tt_exec = 'Commits the plan (enabled after an error-free simulation)'
    btn_newta = 'New TA'; tt_newta = 'Creates a TA inside the selected TA (joins the plan)'
    btn_rename = 'Rename'; tt_rename = 'Renames the selected TA (joins the plan) · F2'
    btn_undo = 'Undo'; tt_undo = 'Removes the last plan item (Ctrl+Z)'
    btn_clear = 'Clear plan'; btn_reload = 'Reload'; tt_reload = 'Reloads the plant from the database (F5)'; btn_opts = 'Options'
    mn_new = 'New TA inside…'; mn_rename = 'Rename or change code…'
    opt_h1 = 'What moves with the TA'
    opt_always = 'Always: jobs, spares, counters (PFO) and job where-lists (ST_WPWHERE2)'
    opt_wh = 'Work history (ST_WorkHistory and ST_WHWHERE2)'; opt_po = 'MR/PO lines (ST_POAITEM2)'
    opt_mat = 'Materials: default TA and transactions'; opt_other = 'Budget, costs and user group restrictions'
    opt_h2 = 'After moving'; opt_retire = 'Deactivate source TAs that end up empty'
    opt_retclass = 'Also deactivate in the function class if no other plant uses it'; opt_items = 'List every moved object in the activity'
    src_title = 'Current structure'; src_sub = 'Drag from here'; src_count = '{0} TAs'
    dst_title = 'New structure'; dst_hint = 'Drop on the parent, even if it is not saved yet'
    chip_new = 'NEW'; chip_react = 'REACTIVATE'; chip_merge = 'MERGE'; chip_rename = 'RENAME'
    plan_title = 'Plan'; plan_sub = 'Runs in order, all or nothing'; tt_remove = 'Remove from plan'
    plan_empty_t = 'Nothing in the plan yet'
    plan_empty_s = 'Drag a TA from the current structure and drop it on the new parent, or right-click to create or rename a TA.'
    act_title = 'Activity'; act_none = 'No simulation yet'; btn_export = 'Export log'
    col_move = 'Item'; col_table = 'Table'; col_action = 'Action'; col_rows = 'Rows'; col_detail = 'Detail (from the database)'
    hist_title = 'History'; hist_sub = 'Every simulation and execution is saved as CSV in the logs folder next to the tool.'
    btn_open = 'Open'; btn_openfolder = 'Open folder'; btn_refresh = 'Refresh'
    col_date = 'Date'; col_kind = 'Type'; col_plant = 'Plant'; col_file = 'File'
    hist_empty = 'No logs yet. Run a simulation on the Relocate screen.'
    conn_title = 'Connection'; conn_sub = 'STAR IPS 3.62 database where the TAs will be relocated.'
    lbl_server = 'Server'; lbl_db = 'Database'; lbl_auth = 'Authentication'; auth_win = 'Windows'; auth_sql = 'SQL login'
    lbl_user = 'User'; lbl_pass = 'Password'; btn_connect = 'Connect'; btn_install = 'Reinstall procedures'
    dlg_title = 'Move structure'; dlg_from = 'FROM'; dlg_to = 'INTO'; dlg_how = 'How should the new codes be applied?'
    rule_prefix = 'Only change the start of the code'; rule_next = 'Use the next free code'; rule_custom = 'Choose the code'
    btn_cancel = 'Cancel'; btn_add = 'Add to plan'; btn_apply = 'Apply'
    col_cur = 'Current code'; col_newc = 'New code'; col_name = 'Name'; col_links = 'Links'; col_status = 'Status'
    dlg_sub = 'Plant {0} · the whole structure moves along ({1} TAs)'
    desc_next = '{0} → {1}, the first free number under {2}'
    desc_next_na = 'A next number cannot be calculated for this code format.'
    desc_custom = 'format {0} · subaccounts follow'
    err_prefix_same = '{0} is already the current parent: changing the start changes nothing. Use the next free code or choose the code.'
    err_next = 'Could not calculate the next free code.'; err_empty = 'Type the new code.'
    err_start = 'The code must start with {0}, the group you dropped on.'
    err_format = 'Use the format {0}: the same level length as {1}.'; err_sep = "After {0} comes '{1}'."
    err_equal = 'The new code is the same as the current one.'; err_inside = 'A TA cannot go inside itself.'
    err_conflict = '{0} TA(s) in conflict: the target is leaving in the plan or does not follow the code pattern.'
    prev = 'Preview'; prev_hdr = 'Preview · {0} TAs'
    cnt_new = '{0} new'; cnt_react = '{0} reactivate'; cnt_merge = '{0} merge'; cnt_bad = '{0} conflict(s)'
    foot = 'Nothing is saved now: the move joins the plan.'
    foot_merge = 'Heads up: {0} TA(s) already exist at the target and will receive the links (merge).'
    rule_t_prefix = 'Change start'; rule_t_next = 'Next free code'; rule_t_custom = 'Chosen code'
    rule_t_create = 'New TA'; rule_t_rename = 'Rename'
    card_create = 'New TA inside {0}'; card_rename = 'was: {0}'
    info_added_t = 'Added to plan'; info_added = '{0} → {1} with {2} TAs. Drag more structures or click Simulate.'
    st_out = 'OFF PATTERN'; st_same = 'NO CHANGE'; st_conflict = 'CONFLICT'; st_merge_plan = 'MERGE (plan)'
    st_merge = 'MERGE'; st_activate = 'ACTIVATE'; st_react = 'REACTIVATE'; st_new = 'NEW'; st_rename = 'RENAME'
    drag_root = 'The root of the chart of accounts cannot be moved.'
    drag_inplan = '{0} is already in the plan. Remove that item to change its target.'
    drag_child = '{0}, which is inside {1}, is already in the plan. Remove it first or move only the rest.'
    drag_dep = 'Plan items use TAs inside {0} as a target or rename them. Remove those items before moving {0}.'
    drag_cant_t = 'This TA cannot be dragged'
    drop_none = 'Drop on the target parent group.'
    drop_leaving = '{0} is leaving in the plan.'; drop_self = 'A TA cannot go inside itself.'
    drag_from = 'Drag a TA from the current structure.'
    dragging = 'Dragging {0} with {1} TA(s): drop on the new parent'; drop_here = 'Drop {0} into {1} · {2}'
    plan_count = '{0} item(s)'; plan_sum = '{0} plan item(s) · {1} TAs'; sum_merge = ' · {0} merge'; sum_bad = ' · {0} conflict(s)'
    plan_emptysum = 'Empty plan'; warn_conf = '{0} conflict(s)'; warn_merge = '{0} merge'
    plan_cascade_t = 'Removed from plan'; plan_cascade = '{0} dependent item(s) were removed too (targets that only existed in the plan).'
    search_t = 'Search'; search_found = "{0} TA(s) found for '{1}'. Esc clears."; search_none = "Nothing found for '{0}'."
    act_line = '{0} · {1} · {2} log lines'; act_sim = 'Simulation'; act_exec = 'Execution'
    resumo = '{0} TAs · {1} PFO links · {2} in ST_WPWHERE2 · {3} in ST_WHWHERE2 · {4} work histories · {5} alerts'
    fail_exec = 'Execution failed, nothing was saved'; fail_sim = 'Simulation failed'
    sim_ok_t = 'Simulation OK, nothing was saved'; sim_ok = '{0}. Review the activity and click Execute.'
    exec_ok_t = 'Plan saved'; exec_ok = '{0}. Log saved to {1}.'
    m_tas = 'TAs in plan'; m_created = 'TAs created / reactivated'; m_pfo = 'Jobs, spares, counters (PFO)'
    m_wp = 'Job where-lists (ST_WPWHERE2)'; m_wh = 'History where-lists (ST_WHWHERE2)'; m_hist = 'Work history (ST_WorkHistory)'
    m_renamed = 'TAs renamed'; m_retired = 'Source TAs deactivated'; m_alerts = 'Alerts'
    log_commit = 'Changes saved.'; log_rollback = 'Simulation: nothing was saved.'; log_err = '{0} (transaction rolled back, nothing saved)'
    log_retire = '{0} (deactivate)'; log_create = '+ {0}'; log_rename = '{0} (rename)'
    proc_missing = "The tool's procedures are missing (or outdated) in {0}.`nInstall them now from usp_IPS_MoveTA.sql?"
    proc_notinst = 'Procedures not installed.'; conn_ok = 'Connected to {0} / {1} (IPS {2}). Procedures ready.'
    conn_fail = 'Failed: {0}'; disconnected = 'Disconnected'; not_connected_t = 'Not connected'
    plant_loaded_t = 'Plant {0} loaded'; plant_loaded = '{0} TAs in class {1}. Drag a TA from the left and drop it on the new parent on the right.'
    plant_fail = 'Failed to load the plant'; connect_first = 'Connect first.'; reinstalled = 'Procedures reinstalled in {0}.'
    switch_plant = 'Switching plants discards the current plan.'
    close_confirm = 'There are plan items that were not executed.'
    log_exported = 'Log exported'; kind_exec = 'EXECUTION'; kind_sim = 'SIMULATION'
    lang_t = 'Language'; lang_set = 'Interface in English. Log details come from the database procedures and stay in Portuguese.'
    cf_exec_title = 'Commit the plan?'; cf_exec_sub = 'Everything runs in a single transaction: all or nothing.'
    cf_tas = 'TAs'; cf_moves = 'Plan items'; cf_plant = 'Plant'; cf_db = 'Database'
    cf_backup = 'If this database is production, take a backup first. The last simulation of this plan passed without errors.'
    cf_go = 'Execute now'; cf_more = '+ {0} more item(s)'
    cf_close_t = 'Close Move TA?'; cf_close_btn = 'Close anyway'
    cf_plant_t = 'Switch plant?'; cf_discard = 'Discard and switch'
    cf_proc_t = 'Install the procedures?'; cf_install = 'Install'
    ed_new_t = 'New Technical Account'; ed_ren_t = 'Rename or change code'
    ed_sub = 'Plant {0} · joins the plan and is only saved when you Execute'
    ed_inside = 'INSIDE'; ed_ta = 'TA'; ed_code = 'Code'; ed_name = 'Name'
    ed_code_hint = 'Suggested: next free code under {0}. You can use another one as long as it starts with {0}.'
    ed_err_code = 'Type the code.'; ed_err_start = 'The code must start with {0} and be longer.'
    ed_err_exists = '{0} already exists in this plant.'; ed_err_planned = '{0} is already in the plan.'
    ed_err_name = 'Type the name.'; ed_err_same = 'The name is the same as the current one.'; ed_err_len = 'Maximum 100 characters.'
    ed_st_new = 'Will be created in class {0} and plant {1}.'
    ed_st_act = 'Already exists in class {0}: it will only be activated in this plant and keeps the class name.'
    ed_st_react = 'Existed deactivated: it will be reactivated with this name.'
    ed_ren_info = 'The name lives in class {0}: it also changes in {1} other plant(s). Filed work history keeps the old name.'
    ed_ren_info0 = 'The name lives in class {0} and on the TA object. Filed work history keeps the old name.'
    ed_ren_ghost = 'This TA is not saved yet: changes apply to how it will be created.'
    ed_foot = 'Nothing is saved now: it joins the plan.'
    ed_cant_moved = '{0} is leaving in the plan. Change its target instead.'
    sel_none_t = 'Select a TA'; sel_none = 'Click a TA in the tree (or right-click it).'
    info_created_t = 'New TA in plan'; info_created = '{0} · {1} goes inside {2}. You can already drop other TAs into it.'
    info_renamed_t = 'Rename in plan'; info_renamed = '{0}: "{1}" → "{2}".'
    ed_code_lock_child = 'The code follows the top of the moved structure ({0}). To change it, edit {0}.'
    ed_code_hint_ren = 'You can change the code too: subaccounts follow. Format {0}.'
    ed_err_same_all = 'Nothing changed.'; ed_err_fmt = 'Use the format {0}.'
    ed_st_recode = 'New code: {0} → {1} · {2} TA(s) change along.'
    st_recode = 'NEW CODE'; rule_t_renumber = 'Renumber'
    info_recoded_t = 'Code changed in plan'; info_recoded = '{0} → {1}: {2} TA(s) follow the new code.'
    rail_about = 'About'; ab_ver = 'Version {0}'; ab_by = 'CREATED BY'; ab_env = 'Environment'; ab_close = 'Close'
    ab_desc = 'Relocates, creates, renames and renumbers Technical Accounts in STAR IPS 3.62, taking jobs, spares, counters, work history and where-lists (ST_WPWHERE2 and ST_WHWHERE2) along. Everything goes through a plan, simulated before it is saved.'
    ab_conn = 'Connection'; ab_ips = 'IPS version'; ab_procs = 'Procedures'; ab_procs_v = '{0} of 4 installed'; ab_logs = 'Log folder'; ab_ps = 'PowerShell / .NET'; ab_none = 'not connected'
}
}
$script:LinkLabelsByLang = @{
    pt = @{ 'jobs'=@('job','jobs'); 'WOs'=@('WO','WOs'); 'WH'=@('hist.','hist.'); 'where-list'=@('where','where'); 'spares'=@('spare','spares'); 'counters'=@('cont.','cont.'); 'equip'=@('equip.','equip.'); 'assemblies'=@('assy','assy'); 'outros'=@('outro','outros') }
    en = @{ 'jobs'=@('job','jobs'); 'WOs'=@('WO','WOs'); 'WH'=@('hist.','hist.'); 'where-list'=@('where','where'); 'spares'=@('spare','spares'); 'counters'=@('counter','counters'); 'equip'=@('equip.','equip.'); 'assemblies'=@('assy','assy'); 'outros'=@('other','others') }
}
$script:ActionEn = @{ 'inicio'='start'; 'fim'='end'; 'existe'='exists'; 'criada'='created'; 'reativada'='reactivated'; 'copiada'='copied'; 'movidas'='moved'; 'atualizadas'='updated'
                      'desativada'='deactivated'; 'renomeada'='renamed'; 'MANTIDA'='KEPT'; 'ATENCAO'='WARNING'; 'VERIFICAR'='CHECK'; 'pulado'='skipped'; 'item'='item'
                      'ja inativa'='already inactive'; 'COMMIT'='COMMIT'; 'ROLLBACK'='ROLLBACK'; 'ERRO'='ERROR' }

function Get-Settings { try { if (Test-Path $script:SettingsFile) { return (Get-Content $script:SettingsFile -Raw | ConvertFrom-Json) } } catch {}; [pscustomobject]@{} }
function Save-Settings([hashtable]$Values) {
    if ($UiTest) { return }
    try {
        $h = @{}; $cur = Get-Settings; foreach ($p in $cur.PSObject.Properties) { $h[$p.Name] = $p.Value }
        foreach ($k in $Values.Keys) { $h[$k] = $Values[$k] }
        $d = Split-Path $script:SettingsFile; if (-not (Test-Path $d)) { New-Item -ItemType Directory $d -Force | Out-Null }
        $h | ConvertTo-Json | Set-Content $script:SettingsFile -Encoding UTF8
    } catch {}
}
function Save-Lang { Save-Settings @{ lang = $script:Lang } }
$script:Saved = Get-Settings
$script:Lang = $(if ($Lang) { $Lang } elseif ($script:Saved.lang -in 'pt','en') { [string]$script:Saved.lang } else { 'pt' })
# sem parâmetro explícito, usa a última conexão que funcionou
if (-not $PSBoundParameters.ContainsKey('Server') -and $script:Saved.server) { $Server = [string]$script:Saved.server }
if (-not $PSBoundParameters.ContainsKey('Database') -and $script:Saved.database) { $Database = [string]$script:Saved.database }
function T([string]$Key) {
    $s = $script:Strings[$script:Lang][$Key]
    if ($null -eq $s) { $s = $script:Strings['pt'][$Key] }
    if ($null -eq $s) { return $Key }
    $s
}

# ============================================================================ modelo (C#)
if (-not ('MoveTA.TaNode' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Collections.ObjectModel;
namespace MoveTA {
  public class Bindable : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    protected void N() { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(string.Empty)); }
  }
  public class TaNode : Bindable {
    public TaNode() { Children = new ObservableCollection<TaNode>(); _state = "normal"; }
    public string Code { get; set; }
    string _name; public string Name { get { return _name; } set { _name = value; N(); } }
    public string Father { get; set; }
    public TaNode Parent { get; set; }
    public ObservableCollection<TaNode> Children { get; private set; }
    string _links, _state, _chip, _chipKind; bool _exp, _drop, _match, _sel;
    public string Links { get { return _links; } set { _links = value; N(); } }
    public string State { get { return _state; } set { _state = value; N(); } }
    public string ChipText { get { return _chip; } set { _chip = value; N(); } }
    public string ChipKind { get { return _chipKind; } set { _chipKind = value; N(); } }
    public bool IsExpanded { get { return _exp; } set { if (_exp != value) { _exp = value; N(); } } }
    public bool IsDropTarget { get { return _drop; } set { if (_drop != value) { _drop = value; N(); } } }
    public bool IsMatch { get { return _match; } set { if (_match != value) { _match = value; N(); } } }
    public bool IsSelected { get { return _sel; } set { if (_sel != value) { _sel = value; N(); } } }
    public bool IsGhost { get { return _state == "new" || _state == "activate" || _state == "reactivate"; } }
    public string RowBg { get { if (_drop) return "#CFE4FA"; if (_match) return "#FFF4CE"; if (IsGhost) return "#EBF3FC"; if (_state == "merge") return "#FFF9F5"; if (_state == "renamed") return "#FAF7FD"; return "Transparent"; } }
    public string RowBorder { get { return _drop ? "#0F6CBD" : "Transparent"; } }
    public string CodeFg { get { if (_state == "moved") return "#9E9E9E"; if (IsGhost) return "#0F548C"; return "#242424"; } }
    public string NameFg { get { if (_state == "moved") return "#9E9E9E"; if (IsGhost) return "#0F548C"; if (_state == "renamed") return "#5C2E91"; return "#424242"; } }
    public string Decor { get { return _state == "moved" ? "Strikethrough" : "None"; } }
    public string LinksVis { get { return string.IsNullOrEmpty(_links) ? "Collapsed" : "Visible"; } }
    public string ChipVis { get { return string.IsNullOrEmpty(_chip) ? "Collapsed" : "Visible"; } }
    public string ChipBg { get { switch (_chipKind) { case "new": return "#0F6CBD"; case "activate": case "reactivate": return "#DFF6DD"; case "merge": return "#FFE8D6"; case "rename": return "#F1E9FA"; case "invalid": return "#FDE7E9"; default: return "#F0F0F0"; } } }
    public string ChipFg { get { switch (_chipKind) { case "new": return "#FFFFFF"; case "activate": case "reactivate": return "#0E700E"; case "merge": return "#8A3707"; case "rename": return "#5C2E91"; case "invalid": return "#B10E1C"; default: return "#616161"; } } }
    public string Tip { get { return Code + "  " + _name + (string.IsNullOrEmpty(_links) ? "" : "\n" + _links); } }
  }
  public class PreviewRow { public string From { get; set; } public string To { get; set; } public string Name { get; set; } public string Links { get; set; } public string Status { get; set; } public string StatusBg { get; set; } public string StatusFg { get; set; } }
  public class PlanCard { public int Index { get; set; } public string From { get; set; } public string To { get; set; } public string ArrowVis { get; set; } public string Name { get; set; } public string Stats { get; set; } public string Rule { get; set; } public string RuleBg { get; set; } public string RuleFg { get; set; } public string Warn { get; set; } public string WarnVis { get; set; } public string WarnBg { get; set; } public string WarnFg { get; set; } }
  public class LogRow { public string Par { get; set; } public string Tabela { get; set; } public string Acao { get; set; } public string AcaoText { get; set; } public string Linhas { get; set; } public string Detalhe { get; set; } public string RowBg { get; set; } }
  public class Metric { public string Key { get; set; } public string Label { get; set; } public string Value { get; set; } public string Fg { get; set; } }
  public class HistRow { public string Date { get; set; } public string Kind { get; set; } public string Plant { get; set; } public string File { get; set; } public string Path { get; set; } public string KindBg { get; set; } public string KindFg { get; set; } }
}
'@
}

# ============================================================================ estado
$script:Conn = $null; $script:ConnLabel = $null; $script:DbName = $null; $script:ConnStatusKey = $null; $script:ConnStatusArgs = @()
$script:PlantId = $null; $script:FuncClass = $null
$script:Base = @{}; $script:ClassSet = @{}; $script:DeletedSet = @{}; $script:Kids = @{}; $script:SortedCodes = @()
$script:Plan = New-Object System.Collections.ArrayList
$script:SrcIndex = @{}; $script:DstIndex = @{}
$script:SimOk = $false; $script:LastRows = $null; $script:LastPairs = 0; $script:LastLogFile = $null; $script:LastKind = $null; $script:LastTime = $null
$script:DlgCtx = $null; $script:EdCtx = $null; $script:DragNode = $null; $script:DragStart = $null
$script:DropNode = $null; $script:HoverNode = $null; $script:HoverSince = [datetime]::MinValue
$script:PendingDrop = $null; $script:Loading = $false; $script:DstScroll = $null
$script:CtxNode = $null; $script:SelNode = $null; $script:MenuItems = @()

function B([string]$Hex) { (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex) }

# ============================================================================ SQL
function New-ConnString { param([string]$Srv, [string]$Db, [bool]$WinAuth, [string]$User, [string]$Pass)
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source'] = $Srv; $b['Initial Catalog'] = $Db; $b['TrustServerCertificate'] = $true
    $b['Application Name'] = 'MoveTA ' + $script:Version; $b['Connect Timeout'] = 8
    if ($WinAuth) { $b['Integrated Security'] = $true } else { $b['User ID'] = $User; $b['Password'] = $Pass }
    $b.ConnectionString
}
function Invoke-Sql {
    param([string]$Sql, [hashtable]$Params = @{}, [System.Data.SqlClient.SqlTransaction]$Tran = $null, [int]$Timeout = 600)
    $cmd = $script:Conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = $Timeout
    if ($Tran) { $cmd.Transaction = $Tran }
    foreach ($k in $Params.Keys) { $v = $Params[$k]; if ($null -eq $v) { $v = [DBNull]::Value }; [void]$cmd.Parameters.AddWithValue($k, $v) }
    $tables = New-Object System.Collections.ArrayList
    $rd = $cmd.ExecuteReader()
    try { do { $t = New-Object System.Data.DataTable; $t.Load($rd); [void]$tables.Add($t) } while (-not $rd.IsClosed -and $rd.HasRows -or ($rd.FieldCount -gt 0 -and -not $rd.IsClosed)) }
    finally { $rd.Close() }
    return ,$tables
}
function Invoke-SqlScalar { param([string]$Sql, [hashtable]$Params = @{})
    $cmd = $script:Conn.CreateCommand(); $cmd.CommandText = $Sql
    foreach ($k in $Params.Keys) { [void]$cmd.Parameters.AddWithValue($k, $Params[$k]) }
    $cmd.ExecuteScalar()
}
function Test-Procs {
    (Invoke-SqlScalar "SELECT CASE WHEN OBJECT_ID('dbo.usp_IPS_MoveTA') IS NOT NULL AND OBJECT_ID('dbo.usp_IPS_RetireTA') IS NOT NULL AND OBJECT_ID('dbo.usp_IPS_CreateTA') IS NOT NULL AND OBJECT_ID('dbo.usp_IPS_RenameTA') IS NOT NULL THEN 1 ELSE 0 END") -eq 1
}
function Install-Procs {
    if (-not (Test-Path $script:ProcFile)) { throw "usp_IPS_MoveTA.sql not found: $($script:ProcFile)" }
    foreach ($b in [regex]::Split([IO.File]::ReadAllText($script:ProcFile), '(?im)^\s*GO\s*$')) {
        if ($b.Trim()) { $cmd = $script:Conn.CreateCommand(); $cmd.CommandText = $b; $cmd.CommandTimeout = 120; [void]$cmd.ExecuteNonQuery() }
    }
}

# ============================================================================ dados da planta
$script:LinkOrder = @('jobs','WOs','WH','where-list','spares','counters','equip','assemblies','outros')
function Format-Links([hashtable]$H) {
    $lab = $script:LinkLabelsByLang[$script:Lang]
    $p = @(); foreach ($k in $script:LinkOrder) { if ($H.ContainsKey($k) -and [int]$H[$k] -gt 0) { $n = [int]$H[$k]; $p += "$n $($lab[$k][[int]($n -ne 1)])" } }
    $p -join ' · '
}
function Update-LinkTexts { foreach ($c in @($script:Base.Keys)) { $script:Base[$c].LinkText = Format-Links $script:Base[$c].Links } }
function Load-Plant([string]$Id) {
    $script:PlantId = $Id
    $script:FuncClass = [string](Invoke-SqlScalar "SELECT FuncClass_Id FROM ST_Plant WHERE Plant_Id = @p" @{ '@p' = $Id })
    $t = (Invoke-Sql @"
SELECT f.Function_Id, f.Function_Father, f.Function_Name,
       CASE WHEN pf.Function_Id IS NULL OR f.Rec_Deleted <> 0 THEN 0 ELSE 1 END AS InPlant,
       CASE WHEN f.Rec_Deleted = 0 THEN 1 ELSE 0 END AS Active
  FROM ST_Function f
  LEFT JOIN ST_PlantFunction pf ON pf.FuncClass_Id = f.FuncClass_Id AND pf.Function_Id = f.Function_Id AND pf.Plant_Id = @p AND pf.Rec_Deleted = 0
 WHERE f.FuncClass_Id = @fc
"@ @{ '@p' = $Id; '@fc' = $script:FuncClass })[0]
    $script:Base = @{}; $script:ClassSet = @{}; $script:DeletedSet = @{}; $script:Kids = @{}
    foreach ($r in $t.Rows) {
        $code = [string]$r.Function_Id
        if ([int]$r.Active -eq 0) { $script:DeletedSet[$code] = 1; continue }
        $script:ClassSet[$code] = 1
        if ([int]$r.InPlant -eq 1) {
            $f = $(if ($r.Function_Father -is [DBNull]) { $null } else { [string]$r.Function_Father })
            $script:Base[$code] = @{ Code = $code; Father = $f; Name = [string]$r.Function_Name; Links = @{}; LinkText = '' }
        }
    }
    foreach ($code in @($script:Base.Keys)) { $f = $script:Base[$code].Father; if ($f -and $script:Base.ContainsKey($f)) { if (-not $script:Kids.ContainsKey($f)) { $script:Kids[$f] = New-Object System.Collections.ArrayList }; [void]$script:Kids[$f].Add($code) } }
    $t2 = (Invoke-Sql @"
SELECT x.Function_Id, x.Kind, COUNT(*) n FROM (
    SELECT pfo.Function_Id, CASE pfo.ObjType_Id WHEN 320 THEN 'jobs' WHEN 340 THEN 'WOs' WHEN 221 THEN 'spares' WHEN 210 THEN 'counters' WHEN 250 THEN 'equip' WHEN 285 THEN 'assemblies' ELSE 'outros' END Kind
      FROM ST_PlantFunctionObject pfo WHERE pfo.Plant_Id = @p AND pfo.FuncClass_Id = @fc AND pfo.Rec_Deleted = 0
    UNION ALL
    SELECT wh.WH_Function_Id, 'WH' FROM ST_WorkHistory wh WHERE wh.Plant_Id = @p AND wh.WH_FuncClass_Id = @fc AND wh.Rec_Deleted = 0
    UNION ALL
    SELECT pf.Function_Id, 'where-list' FROM ST_WPWHERE2 w JOIN ST_PlantFunction pf ON pf.PLANTFUNC_OBJECT_ID = w.OBJECT_ID
     WHERE pf.Plant_Id = @p AND w.REC_DELETED = 0 AND ISNULL(w.WPWHERE_MAIN,0) <> -1
) x GROUP BY x.Function_Id, x.Kind
"@ @{ '@p' = $Id; '@fc' = $script:FuncClass })[0]
    foreach ($r in $t2.Rows) { $c = [string]$r.Function_Id; if ($script:Base.ContainsKey($c)) { $script:Base[$c].Links[[string]$r.Kind] = [int]$r.n } }
    Update-LinkTexts
    $arr = [string[]]@($script:Base.Keys); [Array]::Sort($arr, [StringComparer]::Ordinal); $script:SortedCodes = $arr
}
function Get-Subtree([string]$Code) {
    $out = New-Object System.Collections.Generic.List[string]; $stack = New-Object System.Collections.Stack; $stack.Push($Code)
    while ($stack.Count -gt 0) { $c = [string]$stack.Pop(); $out.Add($c); if ($script:Kids.ContainsKey($c)) { foreach ($k in $script:Kids[$c]) { $stack.Push($k) } } }
    $arr = $out.ToArray(); [Array]::Sort($arr, [StringComparer]::Ordinal); return ,$arr
}
function Test-IsAncestor([string]$Anc, [string]$Code) {
    $c = $Code; $guard = 0
    while ($c -and $guard -lt 50) { if ($c -eq $Anc) { return $true }; if (-not $script:Base.ContainsKey($c)) { return $false }; $c = $script:Base[$c].Father; $guard++ }
    $false
}
function Get-Suffix([string]$Src) {
    $f = $script:Base[$Src].Father
    if ($f -and $f -ne '*' -and $Src.StartsWith($f) -and $Src.Length -gt $f.Length) { return $Src.Substring($f.Length) }
    $Src
}
function Get-NameOf([string]$Code) {
    if ($script:DstIndex.ContainsKey($Code)) { return $script:DstIndex[$Code].Name }
    if ($script:Base.ContainsKey($Code)) { return $script:Base[$Code].Name }
    ''
}

# ============================================================================ plano
# Itens: Type = move   { Src, Target, Rule, NewTop, Pairs, Name, NameOverrides }
#        Type = create { Code, Father, Name, Kind, Status }
#        Type = rename { Code, OldName, NewName }
function Get-Dependency($It) { switch ($It.Type) { 'move' { $It.Target } 'create' { $It.Father } 'rename' { $It.Code } } }
function Get-CreatedCodes($It) {
    $h = @{}
    if ($It.Type -eq 'create') { $h[$It.Code] = 1 }
    elseif ($It.Type -eq 'move') { foreach ($p in $It.Pairs) { if ($p.To -and $p.Kind -in 'new','activate','reactivate') { $h[$p.To] = 1 } } }
    $h
}
function Get-PlannedTargets([int]$Upto = -1) {
    $h = @{}
    for ($i = 0; $i -lt $script:Plan.Count; $i++) {
        if ($Upto -ge 0 -and $i -ge $Upto) { break }
        $it = $script:Plan[$i]
        if ($it.Type -eq 'move') { foreach ($p in $it.Pairs) { if ($p.To) { $h[$p.To] = $p.From } } }
        elseif ($it.Type -eq 'create') { $h[$it.Code] = '+' }
    }
    $h
}
function Get-PlannedSources([int]$Skip = -1) {
    $h = @{}
    for ($i = 0; $i -lt $script:Plan.Count; $i++) { if ($i -eq $Skip) { continue }; $it = $script:Plan[$i]; if ($it.Type -eq 'move') { foreach ($p in $it.Pairs) { $h[$p.From] = $p.To } } }
    $h
}
function Get-KindColors([string]$Kind) {
    switch ($Kind) { 'new' { @('#0F6CBD','#FFFFFF') } 'activate' { @('#DFF6DD','#0E700E') } 'reactivate' { @('#DFF6DD','#0E700E') } 'merge' { @('#FFE8D6','#8A3707') } default { @('#FDE7E9','#B10E1C') } }
}
function Get-Mapping([string]$Src, [string]$Top, [int]$Upto = -1) {
    $planned = Get-PlannedTargets $Upto
    $sources = Get-PlannedSources $Upto
    $rows = New-Object System.Collections.ArrayList
    foreach ($c in (Get-Subtree $Src)) {
        $to = $(if ($c.StartsWith($Src)) { $Top + $c.Substring($Src.Length) } else { $null })
        if (-not $to)                                { $st = @((T 'st_out'), 'invalid') }
        elseif ($to -eq $c)                          { $st = @((T 'st_same'), 'invalid') }
        elseif ($sources.ContainsKey($to))           { $st = @((T 'st_conflict'), 'invalid') }
        elseif ($planned.ContainsKey($to))           { $st = @((T 'st_merge_plan'), 'merge') }
        elseif ($script:Base.ContainsKey($to))       { $st = @((T 'st_merge'), 'merge') }
        elseif ($script:ClassSet.ContainsKey($to))   { $st = @((T 'st_activate'), 'activate') }
        elseif ($script:DeletedSet.ContainsKey($to)) { $st = @((T 'st_react'), 'reactivate') }
        else                                         { $st = @((T 'st_new'), 'new') }
        [void]$rows.Add([pscustomobject]@{ From = $c; To = $to; Name = $script:Base[$c].Name; ToName = $null; Status = $st[0]; Kind = $st[1]; Links = $script:Base[$c].LinkText; LinkH = $script:Base[$c].Links })
    }
    return ,$rows
}
function Get-StatsText($Rows) {
    $h = @{}; foreach ($r in $Rows) { foreach ($k in $r.LinkH.Keys) { $h[$k] = [int]$h[$k] + [int]$r.LinkH[$k] } }
    $n = @($Rows).Count; $s = "$n TA" + $(if ($n -ne 1) { 's' } else { '' })
    $l = Format-Links $h; if ($l) { $s += " · $l" }; $s
}
function Get-NextFree([string]$Tgt, [string]$Sep, [int]$Width) {
    if ($Width -le 0 -or $Width -gt 6) { return $null }
    $max = 0; $pref = $Tgt + $Sep
    $pool = @($script:ClassSet.Keys) + @($script:DeletedSet.Keys) + @((Get-PlannedTargets).Keys)
    foreach ($c in $pool) {
        if ($c.Length -eq $pref.Length + $Width -and $c.StartsWith($pref)) { $rest = $c.Substring($pref.Length); $v = 0; if ([int]::TryParse($rest, [ref]$v) -and $v -gt $max) { $max = $v } }
    }
    $s = ($max + 1).ToString().PadLeft($Width, '0')
    if ($s.Length -gt $Width) { return $null }
    $pref + $s
}
function Get-ChildPattern([string]$Parent) {
    $kids = @()
    if ($script:DstIndex.ContainsKey($Parent)) { $kids = @($script:DstIndex[$Parent].Children | ForEach-Object { $_.Code }) }
    foreach ($k in $kids) { if ($k.StartsWith($Parent) -and $k.Length -gt $Parent.Length) { $suf = $k.Substring($Parent.Length); $sep = [regex]::Match($suf, '^[^0-9A-Za-z]*').Value; return @($sep, ($suf.Length - $sep.Length)) } }
    @('.', 2)
}
function Test-DragSource([string]$Code, [string]$IgnoreRename = '') {
    if ($Code -eq '*' -or -not $script:Base[$Code].Father) { return (T 'drag_root') }
    $src = Get-PlannedSources
    if ($src.ContainsKey($Code)) { return ((T 'drag_inplan') -f $Code) }
    foreach ($it in $script:Plan) { if ($it.Type -eq 'move' -and (Test-IsAncestor $Code $it.Src)) { return ((T 'drag_child') -f $it.Src, $Code) } }
    $sub = @{}; foreach ($c in (Get-Subtree $Code)) { $sub[$c] = 1 }
    foreach ($it in $script:Plan) {
        if ($it.Type -eq 'rename' -and $IgnoreRename -and $it.Code -eq $IgnoreRename) { continue }
        $dep = Get-Dependency $it; if ($dep -and $sub.ContainsKey($dep)) { return ((T 'drag_dep') -f $Code) }
    }
    $null
}
# ---- recodificar: quem "produz" um código no plano, troca de prefixo nos itens dependentes
function Find-PlanOwner([string]$Code) {
    for ($i = 0; $i -lt $script:Plan.Count; $i++) {
        $it = $script:Plan[$i]
        if ($it.Type -eq 'create' -and $it.Code -eq $Code) { return @{ Index = $i; Item = $it; Role = 'create' } }
        if ($it.Type -eq 'move') {
            if ($it.NewTop -eq $Code) { return @{ Index = $i; Item = $it; Role = 'movetop' } }
            if (@($it.Pairs | Where-Object { $_.To -eq $Code }).Count) { return @{ Index = $i; Item = $it; Role = 'movechild' } }
        }
    }
    $null
}
function Get-GhostCount([string]$Code) {
    $n = $script:DstIndex[$Code]; if (-not $n) { return 1 }
    $c = 1; foreach ($k in @($n.Children)) { $c += Get-GhostCount $k.Code }; $c
}
function Set-MoveTop($It, [string]$NewTop) {
    $old = $It.NewTop; $It.NewTop = $NewTop
    $no = @{}
    foreach ($k in @($It.NameOverrides.Keys)) { $nk = $(if ($k.StartsWith($old)) { $NewTop + $k.Substring($old.Length) } else { $k }); $no[$nk] = $It.NameOverrides[$k] }
    $It.NameOverrides = $no
}
function Test-InsideCode([string]$Code, [string]$Prefix) { $Code -and ($Code -eq $Prefix -or ($Code.StartsWith($Prefix) -and $Code.Length -gt $Prefix.Length)) }
function Rename-PlanPrefix([string]$Old, [string]$New, $Skip) {
    foreach ($it in $script:Plan) {
        if ([object]::ReferenceEquals($it, $Skip)) { continue }
        if ($it.Type -eq 'move' -and (Test-InsideCode $it.Target $Old)) {
            $it.Target = $New + $it.Target.Substring($Old.Length)
            if (Test-InsideCode $it.NewTop $Old) { Set-MoveTop $it ($New + $it.NewTop.Substring($Old.Length)) }
        }
        elseif ($it.Type -eq 'create' -and (Test-InsideCode $it.Father $Old)) {
            $it.Father = $New + $it.Father.Substring($Old.Length)
            if (Test-InsideCode $it.Code $Old) { $it.Code = $New + $it.Code.Substring($Old.Length) }
        }
    }
}
function Test-DropTarget([string]$Src, $TgtNode) {
    if (-not $TgtNode) { return (T 'drop_none') }
    if ($TgtNode.State -eq 'moved') { return ((T 'drop_leaving') -f $TgtNode.Code) }
    $n = $TgtNode; while ($n) { if ($n.Code -eq $Src) { return (T 'drop_self') }; $n = $n.Parent }
    if (Test-IsAncestor $Src $TgtNode.Code) { return (T 'drop_self') }
    $null
}
function Remove-PlanItem([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $script:Plan.Count) { return }
    $set = @{}; $set[$Index] = 1; $changed = $true
    while ($changed) {
        $changed = $false
        $created = @{}; foreach ($i in @($set.Keys)) { foreach ($k in (Get-CreatedCodes $script:Plan[$i]).Keys) { $created[$k] = 1 } }
        for ($j = 0; $j -lt $script:Plan.Count; $j++) {
            if ($set.ContainsKey($j)) { continue }
            $dep = Get-Dependency $script:Plan[$j]
            if ($dep -and $created.ContainsKey($dep)) { $set[$j] = 1; $changed = $true }
        }
    }
    foreach ($i in (@($set.Keys) | Sort-Object -Descending)) { $script:Plan.RemoveAt([int]$i) }
    $script:SimOk = $false
    Refresh-Plan
    if ($set.Count -gt 1) { Show-Info 'warn' (T 'plan_cascade_t') ((T 'plan_cascade') -f ($set.Count - 1)) }
}

# ============================================================================ XAML
$script:StylesXaml = @'
    <SolidColorBrush x:Key="Accent" Color="#0F6CBD"/>
    <FontFamily x:Key="Icons">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#EBF3FC"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#242424"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}" Color="#F0F0F0"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#242424"/>
    <ControlTemplate x:Key="BtnTpl" TargetType="ButtonBase">
      <Grid>
        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"/>
        <Border x:Name="Ov" Background="#000000" Opacity="0" CornerRadius="4"/>
        <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Ov" Property="Opacity" Value="0.05"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter TargetName="Ov" Property="Opacity" Value="0.12"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
    <Style x:Key="Subtle" TargetType="Button">
      <Setter Property="Template" Value="{StaticResource BtnTpl}"/>
      <Setter Property="Background" Value="Transparent"/><Setter Property="BorderBrush" Value="Transparent"/><Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Foreground" Value="#242424"/><Setter Property="Height" Value="32"/><Setter Property="Padding" Value="10,0"/>
      <Setter Property="FontSize" Value="13"/><Setter Property="Cursor" Value="Hand"/><Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Outline" TargetType="Button" BasedOn="{StaticResource Subtle}">
      <Setter Property="Background" Value="#FFFFFF"/><Setter Property="BorderBrush" Value="#D1D1D1"/><Setter Property="Padding" Value="14,0"/>
    </Style>
    <Style x:Key="AccentBtn" TargetType="Button" BasedOn="{StaticResource Subtle}">
      <Setter Property="Background" Value="#0F6CBD"/><Setter Property="BorderBrush" Value="#0F6CBD"/><Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="Padding" Value="14,0"/><Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="Rail" TargetType="Button" BasedOn="{StaticResource Subtle}">
      <Setter Property="Width" Value="56"/><Setter Property="Height" Value="52"/><Setter Property="Margin" Value="6,2"/><Setter Property="Padding" Value="0"/>
      <Setter Property="Foreground" Value="#424242"/><Setter Property="FontSize" Value="10"/>
    </Style>
    <ControlTemplate x:Key="CapTpl" TargetType="Button">
      <Border x:Name="Bd" Background="Transparent">
        <ContentPresenter TextElement.FontFamily="{StaticResource Icons}" TextElement.FontSize="10" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{Binding Tag, RelativeSource={RelativeSource TemplatedParent}}"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
    <Style x:Key="Caption" TargetType="Button">
      <Setter Property="Template" Value="{StaticResource CapTpl}"/><Setter Property="Width" Value="46"/><Setter Property="Height" Value="48"/>
      <Setter Property="Foreground" Value="#FFFFFF"/><Setter Property="Tag" Value="#26FFFFFF"/>
      <Setter Property="shell:WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
    </Style>
    <ControlTemplate x:Key="LangTpl" TargetType="Button">
      <Border x:Name="Bd" Background="#26FFFFFF" CornerRadius="4" Padding="9,3">
        <StackPanel Orientation="Horizontal">
          <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE774;" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <ContentPresenter VerticalAlignment="Center"/>
        </StackPanel>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#40FFFFFF"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#FFFFFF"/><Setter Property="CornerRadius" Value="8"/>
      <Setter Property="BorderBrush" Value="#E0E0E0"/><Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="32"/><Setter Property="Padding" Value="6,0"/><Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush" Value="#D1D1D1"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#0F6CBD"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Background" Value="#F5F5F5"/><Setter Property="Foreground" Value="#616161"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TreeViewItem">
      <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}"/>
      <Setter Property="IsSelected" Value="{Binding IsSelected, Mode=TwoWay}"/>
      <Setter Property="Padding" Value="0,1"/><Setter Property="FontSize" Value="13"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="AutoGenerateColumns" Value="False"/><Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="HeadersVisibility" Value="Column"/><Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#F0F0F0"/><Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Background" Value="#FFFFFF"/><Setter Property="RowHeaderWidth" Value="0"/>
      <Setter Property="CanUserAddRows" Value="False"/><Setter Property="CanUserResizeRows" Value="False"/>
      <Setter Property="SelectionMode" Value="Single"/><Setter Property="FontSize" Value="12"/><Setter Property="RowHeight" Value="28"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#FAFAFA"/><Setter Property="Foreground" Value="#616161"/><Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="8,6"/><Setter Property="BorderBrush" Value="#E0E0E0"/><Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#EBF3FC"/><Setter Property="Foreground" Value="#242424"/></Trigger></Style.Triggers>
    </Style>
    <Style x:Key="CellText" TargetType="TextBlock"><Setter Property="Margin" Value="8,0"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextTrimming" Value="CharacterEllipsis"/></Style>
    <Style x:Key="CellCode" TargetType="TextBlock" BasedOn="{StaticResource CellText}"><Setter Property="FontWeight" Value="SemiBold"/></Style>
    <Style x:Key="CellAccent" TargetType="TextBlock" BasedOn="{StaticResource CellText}"><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Foreground" Value="#0F548C"/></Style>
    <Style x:Key="CellMuted" TargetType="TextBlock" BasedOn="{StaticResource CellText}"><Setter Property="Foreground" Value="#616161"/></Style>
    <Style x:Key="RuleCard" TargetType="RadioButton">
      <Setter Property="Cursor" Value="Hand"/><Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="Bd" Background="#FFFFFF" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="6" Padding="12,10">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Grid VerticalAlignment="Top" Margin="0,1,0,0" Width="18" Height="18" HorizontalAlignment="Left">
                  <Ellipse x:Name="Ring" Stroke="#616161" StrokeThickness="1.5"/>
                  <Ellipse x:Name="Dot" Fill="#0F6CBD" Width="9" Height="9" Visibility="Collapsed"/>
                </Grid>
                <ContentPresenter Grid.Column="1"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#FAFAFA"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#0F6CBD"/><Setter TargetName="Bd" Property="Background" Value="#F5F9FD"/>
                <Setter TargetName="Ring" Property="Stroke" Value="#0F6CBD"/><Setter TargetName="Dot" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <HierarchicalDataTemplate x:Key="TaTpl" ItemsSource="{Binding Children}">
      <Border Background="{Binding RowBg}" BorderBrush="{Binding RowBorder}" BorderThickness="1" CornerRadius="4" Padding="6,3" ToolTip="{Binding Tip}">
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="{Binding Code}" FontWeight="SemiBold" Foreground="{Binding CodeFg}" TextDecorations="{Binding Decor}" MinWidth="64"/>
          <TextBlock Text="{Binding Name}" Foreground="{Binding NameFg}" TextDecorations="{Binding Decor}" Margin="8,0,0,0"/>
          <Border Visibility="{Binding LinksVis}" Background="#F3F3F3" CornerRadius="4" Padding="5,0" Margin="8,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="{Binding Links}" FontSize="11" Foreground="#616161"/>
          </Border>
          <Border Visibility="{Binding ChipVis}" Background="{Binding ChipBg}" CornerRadius="4" Padding="6,0" Margin="8,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="{Binding ChipText}" FontSize="11" FontWeight="SemiBold" Foreground="{Binding ChipFg}"/>
          </Border>
        </StackPanel>
      </Border>
    </HierarchicalDataTemplate>
'@

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Move TA" Width="1440" Height="900" MinWidth="1180" MinHeight="700"
        WindowStartupLocation="CenterScreen" Background="#F3F3F3"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#242424"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
  <shell:WindowChrome.WindowChrome>
    <shell:WindowChrome CaptionHeight="48" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
  </shell:WindowChrome.WindowChrome>
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Border x:Name="Root" Background="#F3F3F3" BorderBrush="#D1D1D1" BorderThickness="1">
    <Grid>
      <Grid.RowDefinitions><RowDefinition Height="48"/><RowDefinition Height="*"/></Grid.RowDefinitions>

      <Grid Background="#0F6CBD">
        <Grid.ColumnDefinitions><ColumnDefinition Width="300"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="16,0,0,0">
          <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8DE;" FontSize="18" Foreground="White" VerticalAlignment="Center"/>
          <TextBlock Text="Move TA" Foreground="White" FontWeight="SemiBold" FontSize="15" Margin="10,0,0,0" VerticalAlignment="Center"/>
          <Border Background="#33FFFFFF" CornerRadius="4" Padding="6,1" Margin="8,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="TxtVersion" Foreground="White" FontSize="11"/>
          </Border>
        </StackPanel>
        <Border Grid.Column="1" Width="520" Height="32" Background="White" CornerRadius="6" HorizontalAlignment="Center" shell:WindowChrome.IsHitTestVisibleInChrome="True">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="34"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE721;" Foreground="#616161" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14"/>
            <TextBox x:Name="TbSearch" Grid.Column="1" BorderThickness="0" Background="Transparent" Padding="0" FontSize="14"/>
            <TextBlock x:Name="TxtSearchHint" Grid.Column="1" Text="{DynamicResource search_hint}" Foreground="#8A8A8A" VerticalAlignment="Center" Margin="2,0,0,0" IsHitTestVisible="False"/>
          </Grid>
        </Border>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Background="#26FFFFFF" CornerRadius="12" Padding="10,3" VerticalAlignment="Center" Margin="0,0,8,0">
            <StackPanel Orientation="Horizontal">
              <Ellipse x:Name="DotConn" Width="8" Height="8" Fill="#BDBDBD" VerticalAlignment="Center"/>
              <TextBlock x:Name="TxtConnChip" Foreground="White" FontSize="12" Margin="6,0,0,0"/>
            </StackPanel>
          </Border>
          <Button x:Name="BtnLang" Template="{StaticResource LangTpl}" Foreground="White" FontSize="12" FontWeight="SemiBold" Cursor="Hand" Margin="0,0,12,0" ToolTip="{DynamicResource tt_lang}" shell:WindowChrome.IsHitTestVisibleInChrome="True"/>
        </StackPanel>
        <StackPanel Grid.Column="3" Orientation="Horizontal">
          <Border Width="30" Height="30" CornerRadius="15" Background="White" Margin="0,0,12,0" VerticalAlignment="Center">
            <TextBlock Text="FM" Foreground="#0F6CBD" FontWeight="SemiBold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <Button x:Name="BtnMin" Style="{StaticResource Caption}" Content="&#xE921;" ToolTip="{DynamicResource tt_min}"/>
          <Button x:Name="BtnMax" Style="{StaticResource Caption}" Content="&#xE922;" ToolTip="{DynamicResource tt_max}"/>
          <Button x:Name="BtnClose" Style="{StaticResource Caption}" Tag="#C42B1C" Content="&#xE8BB;" ToolTip="{DynamicResource tt_close}"/>
        </StackPanel>
      </Grid>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions><ColumnDefinition Width="68"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

        <StackPanel Margin="0,8,0,0">
          <Button x:Name="RailMove" Style="{StaticResource Rail}" Tag="move">
            <StackPanel><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8DE;" FontSize="20" HorizontalAlignment="Center"/><TextBlock Text="{DynamicResource rail_move}" HorizontalAlignment="Center" Margin="0,3,0,0"/></StackPanel>
          </Button>
          <Button x:Name="RailHist" Style="{StaticResource Rail}" Tag="hist">
            <StackPanel><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE81C;" FontSize="20" HorizontalAlignment="Center"/><TextBlock Text="{DynamicResource rail_hist}" HorizontalAlignment="Center" Margin="0,3,0,0"/></StackPanel>
          </Button>
          <Button x:Name="RailConn" Style="{StaticResource Rail}" Tag="conn">
            <StackPanel><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE703;" FontSize="20" HorizontalAlignment="Center"/><TextBlock Text="{DynamicResource rail_conn}" HorizontalAlignment="Center" Margin="0,3,0,0"/></StackPanel>
          </Button>
          <Rectangle Height="1" Fill="#E0E0E0" Margin="14,8"/>
          <Button x:Name="RailAbout" Style="{StaticResource Rail}" Tag="about">
            <StackPanel><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE946;" FontSize="20" HorizontalAlignment="Center"/><TextBlock Text="{DynamicResource rail_about}" HorizontalAlignment="Center" Margin="0,3,0,0"/></StackPanel>
          </Button>
        </StackPanel>

        <Grid Grid.Column="1">
          <Grid x:Name="PanelMove" Margin="4,8,12,12">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="220"/><RowDefinition Height="8"/><RowDefinition Height="240" MinHeight="120"/></Grid.RowDefinitions>
            <Grid Margin="8,4,0,10">
              <StackPanel Margin="0,0,320,0">
                <TextBlock Text="{DynamicResource move_title}" FontSize="20" FontWeight="SemiBold"/>
                <TextBlock Text="{DynamicResource move_sub}" Foreground="#616161" FontSize="13" Margin="0,2,0,0" TextWrapping="Wrap"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                <TextBlock Text="{DynamicResource plant}" Foreground="#616161" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="13"/>
                <ComboBox x:Name="CbPlant" Width="260" Height="32" VerticalContentAlignment="Center" FontSize="13"/>
              </StackPanel>
            </Grid>
            <StackPanel Grid.Row="1">
              <Border Style="{StaticResource Card}" Padding="6,4">
                <DockPanel LastChildFill="False">
                  <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                    <Button x:Name="BtnSim" Style="{StaticResource Outline}" ToolTip="{DynamicResource tt_sim}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE768;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_sim}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Button x:Name="BtnExec" Style="{StaticResource AccentBtn}" Margin="6,0,0,0" ToolTip="{DynamicResource tt_exec}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE73E;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_exec}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Rectangle Width="1" Fill="#E0E0E0" Margin="10,6"/>
                    <Button x:Name="BtnNewTA" Style="{StaticResource Subtle}" ToolTip="{DynamicResource tt_newta}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE710;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#0E700E"/><TextBlock Text="{DynamicResource btn_newta}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Button x:Name="BtnRename" Style="{StaticResource Subtle}" ToolTip="{DynamicResource tt_rename}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8AC;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#5C2E91"/><TextBlock Text="{DynamicResource btn_rename}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Rectangle Width="1" Fill="#E0E0E0" Margin="10,6"/>
                    <Button x:Name="BtnUndo" Style="{StaticResource Subtle}" ToolTip="{DynamicResource tt_undo}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE7A7;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_undo}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Button x:Name="BtnClear" Style="{StaticResource Subtle}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE74D;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_clear}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <Button x:Name="BtnReload" Style="{StaticResource Subtle}" ToolTip="{DynamicResource tt_reload}">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE72C;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_reload}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <ToggleButton x:Name="TgOptions" Template="{StaticResource BtnTpl}" Background="Transparent" BorderBrush="Transparent" BorderThickness="1" Height="32" Padding="10,0" FontSize="13" Cursor="Hand">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE713;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_opts}" VerticalAlignment="Center"/><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE70D;" FontSize="10" VerticalAlignment="Center" Margin="6,0,0,0"/></StackPanel>
                    </ToggleButton>
                    <Popup x:Name="PopOptions" PlacementTarget="{Binding ElementName=TgOptions}" Placement="Bottom" StaysOpen="False" AllowsTransparency="True" IsOpen="{Binding IsChecked, ElementName=TgOptions, Mode=TwoWay}">
                      <Border Background="White" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,4,10,10">
                        <Border.Effect><DropShadowEffect BlurRadius="16" ShadowDepth="2" Opacity="0.18"/></Border.Effect>
                        <StackPanel>
                          <TextBlock Text="{DynamicResource opt_h1}" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,6"/>
                          <TextBlock Text="{DynamicResource opt_always}" Foreground="#616161" FontSize="12" Margin="0,0,0,8"/>
                          <CheckBox x:Name="ChkWH" IsChecked="True" Content="{DynamicResource opt_wh}" Margin="0,3"/>
                          <CheckBox x:Name="ChkPO" IsChecked="True" Content="{DynamicResource opt_po}" Margin="0,3"/>
                          <CheckBox x:Name="ChkMat" IsChecked="True" Content="{DynamicResource opt_mat}" Margin="0,3"/>
                          <CheckBox x:Name="ChkOther" IsChecked="True" Content="{DynamicResource opt_other}" Margin="0,3"/>
                          <TextBlock Text="{DynamicResource opt_h2}" FontWeight="SemiBold" FontSize="13" Margin="0,12,0,6"/>
                          <CheckBox x:Name="ChkRetire" IsChecked="True" Content="{DynamicResource opt_retire}" Margin="0,3"/>
                          <CheckBox x:Name="ChkRetireClass" IsChecked="True" Content="{DynamicResource opt_retclass}" Margin="0,3"/>
                          <CheckBox x:Name="ChkItems" IsChecked="False" Content="{DynamicResource opt_items}" Margin="0,3"/>
                        </StackPanel>
                      </Border>
                    </Popup>
                  </StackPanel>
                  <TextBlock x:Name="TxtPlanSummary" DockPanel.Dock="Right" VerticalAlignment="Center" Foreground="#616161" FontSize="12" Margin="0,0,10,0"/>
                </DockPanel>
              </Border>
              <Border x:Name="InfoBar" Visibility="Collapsed" CornerRadius="6" Padding="12,8" Margin="0,8,0,0" BorderThickness="1">
                <DockPanel>
                  <Button x:Name="BtnInfoClose" DockPanel.Dock="Right" Style="{StaticResource Subtle}" Height="24" Padding="6,0" VerticalAlignment="Top">
                    <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8BB;" FontSize="10"/>
                  </Button>
                  <TextBlock x:Name="InfoIcon" FontFamily="{StaticResource Icons}" FontSize="16" VerticalAlignment="Top" Margin="0,1,10,0"/>
                  <TextBlock x:Name="InfoTitle" FontWeight="SemiBold" VerticalAlignment="Top" Margin="0,0,8,0"/>
                  <TextBlock x:Name="InfoText" VerticalAlignment="Top" TextWrapping="Wrap"/>
                </DockPanel>
              </Border>
            </StackPanel>

            <Grid Grid.Row="2" Margin="0,8,0,0">
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="36"/><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="330"/></Grid.ColumnDefinitions>
              <Border Style="{StaticResource Card}">
                <DockPanel>
                  <Border DockPanel.Dock="Top" Padding="14,12,14,8" BorderBrush="#F0F0F0" BorderThickness="0,0,0,1">
                    <DockPanel>
                      <TextBlock x:Name="TxtSrcCount" DockPanel.Dock="Right" Foreground="#616161" FontSize="12" VerticalAlignment="Center"/>
                      <StackPanel>
                        <TextBlock Text="{DynamicResource src_title}" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock Text="{DynamicResource src_sub}" Foreground="#616161" FontSize="12"/>
                      </StackPanel>
                    </DockPanel>
                  </Border>
                  <TreeView x:Name="TreeSrc" BorderThickness="0" Padding="6" ItemTemplate="{StaticResource TaTpl}" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
                </DockPanel>
              </Border>
              <TextBlock Grid.Column="1" FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="18" Foreground="#8A8A8A" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              <Border Grid.Column="2" Style="{StaticResource Card}">
                <DockPanel>
                  <Border DockPanel.Dock="Top" Padding="14,12,14,8" BorderBrush="#F0F0F0" BorderThickness="0,0,0,1">
                    <DockPanel>
                      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
                        <Border Background="#0F6CBD" CornerRadius="4" Padding="6,0" Margin="4,0,0,0"><TextBlock Text="{DynamicResource chip_new}" FontSize="10" FontWeight="SemiBold" Foreground="White"/></Border>
                        <Border Background="#DFF6DD" CornerRadius="4" Padding="6,0" Margin="4,0,0,0"><TextBlock Text="{DynamicResource chip_react}" FontSize="10" FontWeight="SemiBold" Foreground="#0E700E"/></Border>
                        <Border Background="#FFE8D6" CornerRadius="4" Padding="6,0" Margin="4,0,0,0"><TextBlock Text="{DynamicResource chip_merge}" FontSize="10" FontWeight="SemiBold" Foreground="#8A3707"/></Border>
                        <Border Background="#F1E9FA" CornerRadius="4" Padding="6,0" Margin="4,0,0,0"><TextBlock Text="{DynamicResource chip_rename}" FontSize="10" FontWeight="SemiBold" Foreground="#5C2E91"/></Border>
                      </StackPanel>
                      <StackPanel>
                        <TextBlock Text="{DynamicResource dst_title}" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock x:Name="TxtDstHint" Foreground="#616161" FontSize="12" TextTrimming="CharacterEllipsis"/>
                      </StackPanel>
                    </DockPanel>
                  </Border>
                  <TreeView x:Name="TreeDst" BorderThickness="0" Padding="6" AllowDrop="True" ItemTemplate="{StaticResource TaTpl}" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
                </DockPanel>
              </Border>
              <Border Grid.Column="4" Style="{StaticResource Card}">
                <DockPanel>
                  <Border DockPanel.Dock="Top" Padding="14,12,14,8" BorderBrush="#F0F0F0" BorderThickness="0,0,0,1">
                    <DockPanel>
                      <TextBlock x:Name="TxtPlanCount" DockPanel.Dock="Right" Foreground="#616161" FontSize="12" VerticalAlignment="Center"/>
                      <StackPanel>
                        <TextBlock Text="{DynamicResource plan_title}" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock Text="{DynamicResource plan_sub}" Foreground="#616161" FontSize="12"/>
                      </StackPanel>
                    </DockPanel>
                  </Border>
                  <Grid>
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                      <ItemsControl x:Name="PlanList" Margin="10,10,10,2">
                        <ItemsControl.ItemTemplate>
                          <DataTemplate>
                            <Border Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,0,8">
                              <DockPanel>
                                <Button DockPanel.Dock="Right" Tag="{Binding Index}" Style="{StaticResource Subtle}" Height="28" Width="28" Padding="0" ToolTip="{DynamicResource tt_remove}" VerticalAlignment="Top">
                                  <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE74D;" FontSize="12" IsHitTestVisible="False"/>
                                </Button>
                                <StackPanel>
                                  <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="{Binding From}" FontWeight="SemiBold"/>
                                    <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="11" Margin="8,0" VerticalAlignment="Center" Foreground="#616161" Visibility="{Binding ArrowVis}"/>
                                    <TextBlock Text="{Binding To}" FontWeight="SemiBold" Foreground="#0F548C"/>
                                  </StackPanel>
                                  <TextBlock Text="{Binding Name}" Foreground="#424242" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
                                  <TextBlock Text="{Binding Stats}" Foreground="#616161" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                  <WrapPanel Margin="0,6,0,0">
                                    <Border Background="{Binding RuleBg}" CornerRadius="4" Padding="6,1"><TextBlock Text="{Binding Rule}" FontSize="11" Foreground="{Binding RuleFg}"/></Border>
                                    <Border Background="{Binding WarnBg}" CornerRadius="4" Padding="6,1" Margin="6,0,0,0" Visibility="{Binding WarnVis}"><TextBlock Text="{Binding Warn}" FontSize="11" FontWeight="SemiBold" Foreground="{Binding WarnFg}"/></Border>
                                  </WrapPanel>
                                </StackPanel>
                              </DockPanel>
                            </Border>
                          </DataTemplate>
                        </ItemsControl.ItemTemplate>
                      </ItemsControl>
                    </ScrollViewer>
                    <StackPanel x:Name="PlanEmpty" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="28">
                      <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8DE;" FontSize="34" Foreground="#BDBDBD" HorizontalAlignment="Center"/>
                      <TextBlock Text="{DynamicResource plan_empty_t}" FontWeight="SemiBold" HorizontalAlignment="Center" Margin="0,10,0,4"/>
                      <TextBlock Text="{DynamicResource plan_empty_s}" Foreground="#616161" FontSize="12" TextWrapping="Wrap" TextAlignment="Center"/>
                    </StackPanel>
                  </Grid>
                </DockPanel>
              </Border>
            </Grid>

            <GridSplitter Grid.Row="3" Height="8" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
            <Border Grid.Row="4" Style="{StaticResource Card}">
              <DockPanel>
                <Border DockPanel.Dock="Top" Padding="14,10,14,6">
                  <DockPanel>
                    <Button x:Name="BtnSaveLog" DockPanel.Dock="Right" Style="{StaticResource Subtle}" IsEnabled="False">
                      <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE74E;" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_export}" VerticalAlignment="Center"/></StackPanel>
                    </Button>
                    <TextBlock Text="{DynamicResource act_title}" FontSize="16" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtActivity" Foreground="#616161" FontSize="12" Margin="12,0,0,0" VerticalAlignment="Center"/>
                  </DockPanel>
                </Border>
                <ItemsControl x:Name="Metrics" DockPanel.Dock="Top" Margin="10,0,10,4">
                  <ItemsControl.ItemsPanel><ItemsPanelTemplate><WrapPanel/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
                  <ItemsControl.ItemTemplate>
                    <DataTemplate>
                      <Border Background="#FAFAFA" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="6" Padding="12,5" Margin="4" MinWidth="110">
                        <StackPanel><TextBlock Text="{Binding Value}" FontSize="18" FontWeight="SemiBold" Foreground="{Binding Fg}"/><TextBlock Text="{Binding Label}" FontSize="11" Foreground="#616161"/></StackPanel>
                      </Border>
                    </DataTemplate>
                  </ItemsControl.ItemTemplate>
                </ItemsControl>
                <DataGrid x:Name="GridLog" Margin="1,0,1,1">
                  <DataGrid.RowStyle><Style TargetType="DataGridRow"><Setter Property="Background" Value="{Binding RowBg}"/></Style></DataGrid.RowStyle>
                  <DataGrid.Columns>
                    <DataGridTextColumn Binding="{Binding Par}" Width="190" ElementStyle="{StaticResource CellCode}"/>
                    <DataGridTextColumn Binding="{Binding Tabela}" Width="190" ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Binding="{Binding AcaoText}" Width="110" ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Binding="{Binding Linhas}" Width="60" ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Binding="{Binding Detalhe}" Width="*" ElementStyle="{StaticResource CellMuted}"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </Grid>

          <Grid x:Name="PanelHist" Visibility="Collapsed" Margin="4,8,12,12">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <StackPanel Margin="8,4,0,14">
              <TextBlock Text="{DynamicResource hist_title}" FontSize="20" FontWeight="SemiBold"/>
              <TextBlock Text="{DynamicResource hist_sub}" Foreground="#616161" FontSize="13" Margin="0,2,0,0"/>
            </StackPanel>
            <Border Grid.Row="1" Style="{StaticResource Card}">
              <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="10,8">
                  <Button x:Name="BtnOpenLog" Style="{StaticResource Outline}" Content="{DynamicResource btn_open}"/>
                  <Button x:Name="BtnOpenFolder" Style="{StaticResource Subtle}" Content="{DynamicResource btn_openfolder}" Margin="6,0,0,0"/>
                  <Button x:Name="BtnRefreshLogs" Style="{StaticResource Subtle}" Content="{DynamicResource btn_refresh}" Margin="6,0,0,0"/>
                </StackPanel>
                <Grid>
                  <DataGrid x:Name="GridLogs">
                    <DataGrid.Columns>
                      <DataGridTextColumn Binding="{Binding Date}" Width="160" ElementStyle="{StaticResource CellText}"/>
                      <DataGridTemplateColumn Width="120">
                        <DataGridTemplateColumn.CellTemplate><DataTemplate>
                          <Border Background="{Binding KindBg}" CornerRadius="4" Padding="6,1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="8,0"><TextBlock Text="{Binding Kind}" Foreground="{Binding KindFg}" FontSize="11" FontWeight="SemiBold"/></Border>
                        </DataTemplate></DataGridTemplateColumn.CellTemplate>
                      </DataGridTemplateColumn>
                      <DataGridTextColumn Binding="{Binding Plant}" Width="120" ElementStyle="{StaticResource CellText}"/>
                      <DataGridTextColumn Binding="{Binding File}" Width="*" ElementStyle="{StaticResource CellMuted}"/>
                    </DataGrid.Columns>
                  </DataGrid>
                  <TextBlock x:Name="TxtHistEmpty" Text="{DynamicResource hist_empty}" Foreground="#616161" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Grid>
              </DockPanel>
            </Border>
          </Grid>

          <Grid x:Name="PanelConn" Visibility="Collapsed" Margin="4,8,12,12">
            <StackPanel Width="560" HorizontalAlignment="Left" Margin="8,4,0,0">
              <TextBlock Text="{DynamicResource conn_title}" FontSize="20" FontWeight="SemiBold"/>
              <TextBlock Text="{DynamicResource conn_sub}" Foreground="#616161" FontSize="13" Margin="0,2,0,14"/>
              <Border Style="{StaticResource Card}" Padding="20">
                <StackPanel>
                  <TextBlock Text="{DynamicResource lbl_server}" FontSize="12" Foreground="#616161" Margin="0,0,0,4"/>
                  <TextBox x:Name="TbServer"/>
                  <TextBlock Text="{DynamicResource lbl_db}" FontSize="12" Foreground="#616161" Margin="0,12,0,4"/>
                  <TextBox x:Name="TbDb"/>
                  <TextBlock Text="{DynamicResource lbl_auth}" FontSize="12" Foreground="#616161" Margin="0,12,0,4"/>
                  <StackPanel Orientation="Horizontal">
                    <RadioButton x:Name="RbWin" Content="{DynamicResource auth_win}" IsChecked="True" GroupName="auth" Margin="0,0,18,0"/>
                    <RadioButton x:Name="RbSql" Content="{DynamicResource auth_sql}" GroupName="auth"/>
                  </StackPanel>
                  <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <StackPanel><TextBlock Text="{DynamicResource lbl_user}" FontSize="12" Foreground="#616161" Margin="0,0,0,4"/><TextBox x:Name="TbUser" IsEnabled="False"/></StackPanel>
                    <StackPanel Grid.Column="2"><TextBlock Text="{DynamicResource lbl_pass}" FontSize="12" Foreground="#616161" Margin="0,0,0,4"/><PasswordBox x:Name="PbPass" Height="32" VerticalContentAlignment="Center" IsEnabled="False" BorderBrush="#D1D1D1"/></StackPanel>
                  </Grid>
                  <StackPanel Orientation="Horizontal" Margin="0,18,0,0">
                    <Button x:Name="BtnConnect" Style="{StaticResource AccentBtn}" Content="{DynamicResource btn_connect}" MinWidth="120"/>
                    <Button x:Name="BtnInstall" Style="{StaticResource Outline}" Content="{DynamicResource btn_install}" Margin="8,0,0,0"/>
                  </StackPanel>
                  <TextBlock x:Name="TxtConnStatus" Margin="0,12,0,0" TextWrapping="Wrap" Foreground="#616161" FontSize="12"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </Grid>
      </Grid>
    </Grid>
  </Border>
</Window>
'@

$script:DialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Move TA" Width="780" Height="720" WindowStyle="None" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="White"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#242424" UseLayoutRounding="True">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Border x:Name="DlgRoot" BorderBrush="#BDBDBD" BorderThickness="1" Background="White">
    <Grid>
      <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
      <Border x:Name="DlgHeader" Padding="24,18,12,4" Background="White">
        <DockPanel>
          <Button x:Name="BtnX" DockPanel.Dock="Right" Style="{StaticResource Subtle}" Width="32" Padding="0" VerticalAlignment="Top">
            <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8BB;" FontSize="11"/>
          </Button>
          <StackPanel>
            <TextBlock Text="{DynamicResource dlg_title}" FontSize="20" FontWeight="SemiBold"/>
            <TextBlock x:Name="TxtSub" Foreground="#616161" FontSize="13" Margin="0,2,0,0"/>
          </StackPanel>
        </DockPanel>
      </Border>
      <Grid Grid.Row="1" Margin="24,10,24,0">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <Border Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="8" Padding="16,12">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="48"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <StackPanel>
              <TextBlock Text="{DynamicResource dlg_from}" FontSize="11" Foreground="#616161" FontWeight="SemiBold"/>
              <TextBlock x:Name="TxtSrcCode" FontSize="20" FontWeight="SemiBold"/>
              <TextBlock x:Name="TxtSrcName" Foreground="#424242" FontSize="12" TextTrimming="CharacterEllipsis"/>
              <TextBlock x:Name="TxtStats" Foreground="#616161" FontSize="12" Margin="0,6,0,0" TextWrapping="Wrap"/>
            </StackPanel>
            <TextBlock Grid.Column="1" FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="20" Foreground="#0F6CBD" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            <StackPanel Grid.Column="2">
              <TextBlock Text="{DynamicResource dlg_to}" FontSize="11" Foreground="#616161" FontWeight="SemiBold"/>
              <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="TxtTgtCode" FontSize="20" FontWeight="SemiBold" Foreground="#0F548C"/>
                <Border x:Name="ChipTgtNew" Background="#0F6CBD" CornerRadius="4" Padding="6,0" Margin="8,0,0,0" VerticalAlignment="Center" Visibility="Collapsed"><TextBlock Text="{DynamicResource chip_new}" FontSize="10" FontWeight="SemiBold" Foreground="White"/></Border>
              </StackPanel>
              <TextBlock x:Name="TxtTgtName" Foreground="#424242" FontSize="12" TextTrimming="CharacterEllipsis"/>
            </StackPanel>
          </Grid>
        </Border>
        <TextBlock Grid.Row="1" Text="{DynamicResource dlg_how}" FontWeight="SemiBold" Margin="0,16,0,8"/>
        <StackPanel Grid.Row="2">
          <RadioButton x:Name="RbPrefix" GroupName="rule" Style="{StaticResource RuleCard}">
            <StackPanel><TextBlock Text="{DynamicResource rule_prefix}" FontWeight="SemiBold"/><TextBlock x:Name="TxtPrefixDesc" Foreground="#616161" FontSize="12" Margin="0,2,0,0" TextWrapping="Wrap"/></StackPanel>
          </RadioButton>
          <RadioButton x:Name="RbNext" GroupName="rule" Style="{StaticResource RuleCard}">
            <StackPanel><TextBlock Text="{DynamicResource rule_next}" FontWeight="SemiBold"/><TextBlock x:Name="TxtNextDesc" Foreground="#616161" FontSize="12" Margin="0,2,0,0" TextWrapping="Wrap"/></StackPanel>
          </RadioButton>
          <RadioButton x:Name="RbCustom" GroupName="rule" Style="{StaticResource RuleCard}">
            <StackPanel>
              <TextBlock Text="{DynamicResource rule_custom}" FontWeight="SemiBold"/>
              <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                <TextBox x:Name="TbCustom" Width="170" FontSize="13"/>
                <TextBlock x:Name="TxtCustomDesc" Foreground="#616161" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
            </StackPanel>
          </RadioButton>
        </StackPanel>
        <DockPanel Grid.Row="3" Margin="0,8,0,6">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
            <Border x:Name="ChipNew" Background="#0F6CBD" CornerRadius="4" Padding="6,1" Margin="4,0,0,0"><TextBlock x:Name="CntNew" FontSize="11" FontWeight="SemiBold" Foreground="White"/></Border>
            <Border x:Name="ChipReact" Background="#DFF6DD" CornerRadius="4" Padding="6,1" Margin="4,0,0,0"><TextBlock x:Name="CntReact" FontSize="11" FontWeight="SemiBold" Foreground="#0E700E"/></Border>
            <Border x:Name="ChipMerge" Background="#FFE8D6" CornerRadius="4" Padding="6,1" Margin="4,0,0,0"><TextBlock x:Name="CntMerge" FontSize="11" FontWeight="SemiBold" Foreground="#8A3707"/></Border>
            <Border x:Name="ChipBad" Background="#FDE7E9" CornerRadius="4" Padding="6,1" Margin="4,0,0,0"><TextBlock x:Name="CntBad" FontSize="11" FontWeight="SemiBold" Foreground="#B10E1C"/></Border>
          </StackPanel>
          <TextBlock x:Name="TxtPrevHdr" FontWeight="SemiBold" VerticalAlignment="Center"/>
        </DockPanel>
        <Border Grid.Row="4" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="6">
          <DataGrid x:Name="GridPrev" Margin="1">
            <DataGrid.Columns>
              <DataGridTextColumn Binding="{Binding From}" Width="120" ElementStyle="{StaticResource CellCode}"/>
              <DataGridTextColumn Binding="{Binding To}" Width="120" ElementStyle="{StaticResource CellAccent}"/>
              <DataGridTextColumn Binding="{Binding Name}" Width="230" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Binding="{Binding Links}" Width="110" ElementStyle="{StaticResource CellMuted}"/>
              <DataGridTemplateColumn Width="118">
                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                  <Border Background="{Binding StatusBg}" CornerRadius="4" Padding="6,1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="8,0"><TextBlock Text="{Binding Status}" Foreground="{Binding StatusFg}" FontSize="11" FontWeight="SemiBold"/></Border>
                </DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
            </DataGrid.Columns>
          </DataGrid>
        </Border>
        <TextBlock Grid.Row="5" x:Name="TxtErr" Foreground="#B10E1C" Margin="0,8,0,0" TextWrapping="Wrap" FontSize="13"/>
      </Grid>
      <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" Padding="24,14" Margin="0,12,0,0">
        <DockPanel>
          <TextBlock x:Name="TxtFoot" DockPanel.Dock="Left" Foreground="#616161" FontSize="12" VerticalAlignment="Center" TextWrapping="Wrap" MaxWidth="400"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnCancel" Style="{StaticResource Outline}" Content="{DynamicResource btn_cancel}" MinWidth="110"/>
            <Button x:Name="BtnOk" Style="{StaticResource AccentBtn}" Margin="8,0,0,0">
              <StackPanel Orientation="Horizontal"><TextBlock FontFamily="{StaticResource Icons}" Text="&#xE710;" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{DynamicResource btn_add}" VerticalAlignment="Center"/></StackPanel>
            </Button>
          </StackPanel>
        </DockPanel>
      </Border>
    </Grid>
  </Border>
</Window>
'@

$script:ConfirmXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Move TA" Width="600" SizeToContent="Height" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#242424" UseLayoutRounding="True">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Grid x:Name="CfRoot" Background="Transparent">
    <Border Margin="18" Background="White" CornerRadius="8" BorderBrush="#D1D1D1" BorderThickness="1">
      <Border.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="4" Opacity="0.22"/></Border.Effect>
      <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <DockPanel x:Name="CfHeader" Margin="24,22,14,0" Background="Transparent">
          <Button x:Name="CfX" DockPanel.Dock="Right" Style="{StaticResource Subtle}" Width="32" Padding="0" VerticalAlignment="Top">
            <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8BB;" FontSize="11"/>
          </Button>
          <Border x:Name="CfIconBg" Width="44" Height="44" CornerRadius="22" VerticalAlignment="Top">
            <TextBlock x:Name="CfIcon" FontFamily="{StaticResource Icons}" FontSize="20" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel Margin="16,2,0,0">
            <TextBlock x:Name="CfTitle" FontSize="20" FontWeight="SemiBold" TextWrapping="Wrap"/>
            <TextBlock x:Name="CfSub" Foreground="#616161" FontSize="13" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </DockPanel>
        <StackPanel Grid.Row="1" Margin="24,18,24,0">
          <ItemsControl x:Name="CfStats">
            <ItemsControl.ItemsPanel><ItemsPanelTemplate><UniformGrid Rows="1"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Border Background="#FAFAFA" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,0,8,0">
                  <StackPanel>
                    <TextBlock Text="{Binding Value}" FontSize="18" FontWeight="SemiBold" Foreground="{Binding Fg}" TextTrimming="CharacterEllipsis" ToolTip="{Binding Value}"/>
                    <TextBlock Text="{Binding Label}" FontSize="11" Foreground="#616161"/>
                  </StackPanel>
                </Border>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
          <Border x:Name="CfMovesBox" Background="#FFFFFF" CornerRadius="6" BorderBrush="#E0E0E0" BorderThickness="1" Margin="0,12,0,0" Padding="14,8">
            <StackPanel>
              <ItemsControl x:Name="CfMoves">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <DockPanel Margin="0,4">
                      <Border DockPanel.Dock="Right" Background="{Binding RuleBg}" CornerRadius="4" Padding="6,1" Margin="12,0,0,0" VerticalAlignment="Center"><TextBlock Text="{Binding Rule}" FontSize="11" Foreground="{Binding RuleFg}"/></Border>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding From}" FontWeight="SemiBold" FontSize="13"/>
                        <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="11" Margin="8,0" VerticalAlignment="Center" Foreground="#616161" Visibility="{Binding ArrowVis}"/>
                        <TextBlock Text="{Binding To}" FontWeight="SemiBold" FontSize="13" Foreground="#0F548C"/>
                        <TextBlock Text="{Binding Name}" Foreground="#616161" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" MaxWidth="250"/>
                      </StackPanel>
                    </DockPanel>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>
              <TextBlock x:Name="CfMore" Foreground="#616161" FontSize="12" Margin="0,4,0,0"/>
            </StackPanel>
          </Border>
          <Border x:Name="CfWarn" Background="#FFF9F5" BorderBrush="#FDCFB4" BorderThickness="1" CornerRadius="6" Padding="12,9" Margin="0,12,0,0">
            <DockPanel>
              <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE7BA;" Foreground="#8A3707" FontSize="15" Margin="0,1,10,0" VerticalAlignment="Top"/>
              <TextBlock x:Name="CfWarnText" TextWrapping="Wrap" FontSize="13" Foreground="#5C2500"/>
            </DockPanel>
          </Border>
        </StackPanel>
        <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" CornerRadius="0,0,8,8" Padding="24,14" Margin="0,22,0,0">
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="CfCancel" Style="{StaticResource Outline}" Content="{DynamicResource btn_cancel}" MinWidth="110"/>
            <Button x:Name="CfOk" Style="{StaticResource AccentBtn}" Margin="8,0,0,0" MinWidth="150">
              <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="CfOkIcon" FontFamily="{StaticResource Icons}" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBlock x:Name="CfOkText" VerticalAlignment="Center"/>
              </StackPanel>
            </Button>
          </StackPanel>
        </Border>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$script:EditXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Move TA" Width="600" SizeToContent="Height" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#242424" UseLayoutRounding="True">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Grid x:Name="EdRoot" Background="Transparent">
    <Border Margin="18" Background="White" CornerRadius="8" BorderBrush="#D1D1D1" BorderThickness="1">
      <Border.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="4" Opacity="0.22"/></Border.Effect>
      <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <DockPanel x:Name="EdHeader" Margin="24,22,14,0" Background="Transparent">
          <Button x:Name="EdX" DockPanel.Dock="Right" Style="{StaticResource Subtle}" Width="32" Padding="0" VerticalAlignment="Top">
            <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8BB;" FontSize="11"/>
          </Button>
          <Border x:Name="EdIconBg" Width="44" Height="44" CornerRadius="22" VerticalAlignment="Top">
            <TextBlock x:Name="EdIcon" FontFamily="{StaticResource Icons}" FontSize="20" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel Margin="16,2,0,0">
            <TextBlock x:Name="EdTitle" FontSize="20" FontWeight="SemiBold"/>
            <TextBlock x:Name="EdSub" Foreground="#616161" FontSize="13" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </DockPanel>
        <StackPanel Grid.Row="1" Margin="24,18,24,0">
          <Border Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="8" Padding="16,10">
            <StackPanel>
              <TextBlock x:Name="EdCtxLabel" FontSize="11" Foreground="#616161" FontWeight="SemiBold"/>
              <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="EdCtxCode" FontSize="18" FontWeight="SemiBold" Foreground="#0F548C"/>
                <TextBlock x:Name="EdCtxName" FontSize="13" Foreground="#424242" Margin="10,0,0,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" MaxWidth="380"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <Grid Margin="0,16,0,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <StackPanel>
              <TextBlock Text="{DynamicResource ed_code}" FontSize="12" Foreground="#616161" Margin="0,0,0,4"/>
              <TextBox x:Name="EdCode" FontSize="14"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock Text="{DynamicResource ed_name}" FontSize="12" Foreground="#616161" Margin="0,0,0,4"/>
              <TextBox x:Name="EdName" FontSize="14" MaxLength="100"/>
            </StackPanel>
          </Grid>
          <TextBlock x:Name="EdHint" Foreground="#616161" FontSize="12" Margin="0,6,0,0" TextWrapping="Wrap"/>
          <Border x:Name="EdStatus" CornerRadius="6" Padding="12,8" Margin="0,12,0,0" BorderThickness="1">
            <DockPanel>
              <Border x:Name="EdChip" CornerRadius="4" Padding="6,1" Margin="0,0,10,0" VerticalAlignment="Top"><TextBlock x:Name="EdChipText" FontSize="11" FontWeight="SemiBold"/></Border>
              <TextBlock x:Name="EdStatusText" FontSize="13" TextWrapping="Wrap"/>
            </DockPanel>
          </Border>
          <TextBlock x:Name="EdErr" Foreground="#B10E1C" Margin="0,8,0,0" TextWrapping="Wrap" FontSize="13"/>
        </StackPanel>
        <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" CornerRadius="0,0,8,8" Padding="24,14" Margin="0,16,0,0">
          <DockPanel>
            <TextBlock DockPanel.Dock="Left" Text="{DynamicResource ed_foot}" Foreground="#616161" FontSize="12" VerticalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
              <Button x:Name="EdCancel" Style="{StaticResource Outline}" Content="{DynamicResource btn_cancel}" MinWidth="110"/>
              <Button x:Name="EdOk" Style="{StaticResource AccentBtn}" Margin="8,0,0,0" MinWidth="150">
                <StackPanel Orientation="Horizontal">
                  <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE710;" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock x:Name="EdOkText" VerticalAlignment="Center"/>
                </StackPanel>
              </Button>
            </StackPanel>
          </DockPanel>
        </Border>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$script:AboutXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Move TA" Width="560" SizeToContent="Height" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#242424" UseLayoutRounding="True">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Grid x:Name="AbRoot" Background="Transparent">
    <Border Margin="18" Background="White" CornerRadius="8" BorderBrush="#D1D1D1" BorderThickness="1">
      <Border.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="4" Opacity="0.22"/></Border.Effect>
      <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <!-- faixa azul com o ícone -->
        <Border x:Name="AbHeader" Background="#0F6CBD" CornerRadius="8,8,0,0" Padding="24,22,14,22">
          <DockPanel>
            <Button x:Name="AbX" DockPanel.Dock="Right" Style="{StaticResource Caption}" Width="32" Height="32" VerticalAlignment="Top" Content="&#xE8BB;"/>
            <Border Width="64" Height="64" CornerRadius="14" Background="White" VerticalAlignment="Center">
              <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE8DE;" FontSize="34" Foreground="#0F6CBD" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Margin="18,0,0,0" VerticalAlignment="Center">
              <TextBlock Text="Move TA" FontSize="26" FontWeight="SemiBold" Foreground="White"/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                <Border Background="#33FFFFFF" CornerRadius="4" Padding="7,1"><TextBlock x:Name="AbVer" Foreground="White" FontSize="12"/></Border>
                <TextBlock Text="STAR IPS 3.62" Foreground="#DDEBFA" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
            </StackPanel>
          </DockPanel>
        </Border>
        <StackPanel Grid.Row="1" Margin="24,18,24,0">
          <TextBlock Text="{DynamicResource ab_desc}" TextWrapping="Wrap" Foreground="#424242" FontSize="13" LineHeight="20"/>
          <Border Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,16,0,0">
            <DockPanel>
              <Border Width="44" Height="44" CornerRadius="22" Background="#0F6CBD" VerticalAlignment="Center">
                <TextBlock Text="FM" Foreground="White" FontWeight="SemiBold" FontSize="15" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Margin="14,0,0,0" VerticalAlignment="Center">
                <TextBlock Text="{DynamicResource ab_by}" FontSize="11" FontWeight="SemiBold" Foreground="#616161"/>
                <TextBlock Text="Flávio Miguel" FontSize="18" FontWeight="SemiBold"/>
              </StackPanel>
            </DockPanel>
          </Border>
          <TextBlock Text="{DynamicResource ab_env}" FontWeight="SemiBold" FontSize="13" Margin="0,16,0,6"/>
          <ItemsControl x:Name="AbEnv">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Grid Margin="0,3">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock Text="{Binding Label}" Foreground="#616161" FontSize="12"/>
                  <TextBlock Grid.Column="1" Text="{Binding Value}" Foreground="{Binding Fg}" FontSize="12" TextWrapping="Wrap"/>
                </Grid>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </StackPanel>
        <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" CornerRadius="0,0,8,8" Padding="24,14" Margin="0,18,0,0">
          <DockPanel>
            <TextBlock x:Name="AbCopy" DockPanel.Dock="Left" Foreground="#616161" FontSize="12" VerticalAlignment="Center"/>
            <Button x:Name="AbClose" Style="{StaticResource AccentBtn}" Content="{DynamicResource ab_close}" MinWidth="110" HorizontalAlignment="Right"/>
          </DockPanel>
        </Border>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

function Show-About([switch]$NoShow) {
    $w = New-Window $script:AboutXaml
    $c = @{}; foreach ($n in @('AbRoot','AbHeader','AbX','AbVer','AbEnv','AbCopy','AbClose')) { $c[$n] = $w.FindName($n) }
    $c.AbVer.Text = (T 'ab_ver') -f $script:Version
    $c.AbCopy.Text = "© $((Get-Date).Year) Flávio Miguel"
    $procs = 0
    if ($script:Conn) { try { $procs = [int](Invoke-SqlScalar "SELECT COUNT(*) FROM sys.procedures WHERE name IN ('usp_IPS_MoveTA','usp_IPS_RetireTA','usp_IPS_CreateTA','usp_IPS_RenameTA')") } catch {} }
    $rows = New-Object 'System.Collections.Generic.List[MoveTA.Metric]'
    foreach ($r in @(
        @((T 'ab_conn'), $(if ($script:Conn) { $script:ConnLabel } else { T 'ab_none' }), $(if ($script:Conn) { '#242424' } else { '#8A3707' })),
        @((T 'ab_ips'), $(if ($script:IpsVer) { $script:IpsVer } else { '-' }), '#242424'),
        @((T 'ab_procs'), $(if ($script:Conn) { (T 'ab_procs_v') -f $procs } else { '-' }), $(if ($procs -eq 4 -or -not $script:Conn) { '#242424' } else { '#B10E1C' })),
        @((T 'ab_logs'), $script:LogDir, '#242424'),
        @((T 'ab_ps'), "$($PSVersionTable.PSVersion) / .NET $([Environment]::Version)", '#242424'))) {
        $x = New-Object MoveTA.Metric; $x.Label = $r[0]; $x.Value = [string]$r[1]; $x.Fg = $r[2]; $rows.Add($x)
    }
    $c.AbEnv.ItemsSource = $rows
    $script:AbWin = $w
    $c.AbClose.Add_Click({ $script:AbWin.Close() }); $c.AbX.Add_Click({ $script:AbWin.Close() })
    $c.AbHeader.Add_MouseLeftButtonDown({ try { $script:AbWin.DragMove() } catch {} })
    $w.Add_KeyDown({ param($s, $e) if ($e.Key -in 'Escape', 'Return') { $script:AbWin.Close() } })
    if ($NoShow) { return @{ Window = $w; Root = $c.AbRoot } }
    if ($script:Win.IsLoaded) { $w.Owner = $script:Win } else { $w.WindowStartupLocation = 'CenterScreen' }
    [void]$w.ShowDialog()
}

function Set-LangResources($Element) { foreach ($k in $script:Strings[$script:Lang].Keys) { $Element.Resources[$k] = [string]$script:Strings[$script:Lang][$k] } }
function Set-ColumnHeaders($Grid, [string[]]$Keys) { for ($i = 0; $i -lt $Keys.Count -and $i -lt $Grid.Columns.Count; $i++) { $Grid.Columns[$i].Header = T $Keys[$i] } }
function New-Window([string]$Xaml) { $w = [Windows.Markup.XamlReader]::Parse($Xaml.Replace('__STYLES__', $script:StylesXaml)); Set-LangResources $w; $w }

# ============================================================================ janela
$win = [Windows.Markup.XamlReader]::Parse($script:MainXaml.Replace('__STYLES__', $script:StylesXaml))
$script:Win = $win
$script:U = @{}
foreach ($n in @('Root','TxtVersion','TbSearch','TxtSearchHint','DotConn','TxtConnChip','BtnLang','BtnMin','BtnMax','BtnClose','RailMove','RailHist','RailConn','RailAbout',
                 'PanelMove','PanelHist','PanelConn','CbPlant','BtnSim','BtnExec','BtnNewTA','BtnRename','BtnUndo','BtnClear','BtnReload','TgOptions','PopOptions',
                 'ChkWH','ChkPO','ChkMat','ChkOther','ChkRetire','ChkRetireClass','ChkItems','TxtPlanSummary','InfoBar','InfoIcon','InfoTitle','InfoText','BtnInfoClose',
                 'TreeSrc','TxtSrcCount','TreeDst','TxtDstHint','PlanList','PlanEmpty','TxtPlanCount','BtnSaveLog','TxtActivity','Metrics','GridLog',
                 'TbServer','TbDb','RbWin','RbSql','TbUser','PbPass','BtnConnect','BtnInstall','TxtConnStatus','GridLogs','BtnOpenLog','BtnOpenFolder','BtnRefreshLogs','TxtHistEmpty')) {
    $script:U[$n] = $win.FindName($n)
}
$U = $script:U
$U.TxtVersion.Text = "v$($script:Version)"; $win.Title = "Move TA $($script:Version)"
$U.TbServer.Text = $Server; $U.TbDb.Text = $Database

# ============================================================================ UI helpers
function Show-Info([string]$Kind, [string]$Title, [string]$Text) {
    $m = @{ info = @('#EBF3FC','#B4D6FA',0xE946,'#0F548C'); success = @('#F1FAF1','#9FD89F',0xE930,'#0E700E'); warn = @('#FFF9F5','#FDCFB4',0xE7BA,'#8A3707'); error = @('#FDF3F4','#EEACB2',0xE783,'#B10E1C') }[$Kind]
    $script:U.InfoBar.Background = B $m[0]; $script:U.InfoBar.BorderBrush = B $m[1]
    $script:U.InfoIcon.Text = [string][char]$m[2]; $script:U.InfoIcon.Foreground = B $m[3]
    $script:U.InfoTitle.Text = $Title; $script:U.InfoText.Text = $Text
    $script:U.InfoBar.Visibility = 'Visible'
}
function Show-Confirm {
    param([string]$Title, [string]$Sub, [string]$OkText, [ValidateSet('warn','question')][string]$Kind = 'question',
          $Stats = $null, $Moves = $null, [string]$More = '', [string]$Warn = '', [string]$OkGlyph = '', [switch]$NoShow)
    $w = New-Window $script:ConfirmXaml
    $c = @{}; foreach ($n in @('CfRoot','CfHeader','CfX','CfIconBg','CfIcon','CfTitle','CfSub','CfStats','CfMovesBox','CfMoves','CfMore','CfWarn','CfWarnText','CfCancel','CfOk','CfOkIcon','CfOkText')) { $c[$n] = $w.FindName($n) }
    if ($Kind -eq 'warn') { $c.CfIconBg.Background = B '#FFF4CE'; $c.CfIcon.Foreground = B '#8A3707'; $c.CfIcon.Text = [string][char]0xE7BA }
    else { $c.CfIconBg.Background = B '#EBF3FC'; $c.CfIcon.Foreground = B '#0F548C'; $c.CfIcon.Text = [string][char]0xE9CE }
    $c.CfTitle.Text = $Title; $c.CfSub.Text = $Sub
    $c.CfSub.Visibility = $(if ($Sub) { 'Visible' } else { 'Collapsed' })
    if ($Stats) { $c.CfStats.ItemsSource = $Stats } else { $c.CfStats.Visibility = 'Collapsed' }
    if ($Moves) { $c.CfMoves.ItemsSource = $Moves; $c.CfMore.Text = $More; $c.CfMore.Visibility = $(if ($More) { 'Visible' } else { 'Collapsed' }) } else { $c.CfMovesBox.Visibility = 'Collapsed' }
    if ($Warn) { $c.CfWarnText.Text = $Warn } else { $c.CfWarn.Visibility = 'Collapsed' }
    $c.CfOkText.Text = $OkText
    if ($OkGlyph) { $c.CfOkIcon.Text = $OkGlyph } else { $c.CfOkIcon.Visibility = 'Collapsed' }
    $script:CfWin = $w
    $c.CfOk.Add_Click({ $script:CfWin.DialogResult = $true })
    $c.CfCancel.Add_Click({ $script:CfWin.DialogResult = $false })
    $c.CfX.Add_Click({ $script:CfWin.DialogResult = $false })
    $c.CfHeader.Add_MouseLeftButtonDown({ try { $script:CfWin.DragMove() } catch {} })
    $w.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'Escape') { $script:CfWin.DialogResult = $false } })
    if ($NoShow) { return @{ Window = $w; Root = $c.CfRoot } }
    if ($script:Win.IsLoaded) { $w.Owner = $script:Win } else { $w.WindowStartupLocation = 'CenterScreen' }
    $c.CfCancel.Focus() | Out-Null
    [bool]$w.ShowDialog()
}
function Get-PlanCards([int]$Max = 1000) {
    $cards = New-Object 'System.Collections.Generic.List[MoveTA.PlanCard]'
    for ($i = 0; $i -lt [Math]::Min($Max, $script:Plan.Count); $i++) {
        $it = $script:Plan[$i]; $card = New-Object MoveTA.PlanCard; $card.Index = $i; $card.WarnVis = 'Collapsed'; $card.ArrowVis = 'Visible'
        switch ($it.Type) {
            'move' {
                $m = @($it.Pairs | Where-Object { $_.Kind -eq 'merge' }).Count; $x = @($it.Pairs | Where-Object { $_.Kind -eq 'invalid' }).Count
                $card.From = $it.Src; $card.To = $it.NewTop; $card.Name = $it.Name; $card.Stats = Get-StatsText $it.Pairs
                $card.Rule = $it.RuleText; $card.RuleBg = '#EBF3FC'; $card.RuleFg = '#0F548C'
                if ($x -gt 0) { $card.Warn = (T 'warn_conf') -f $x; $card.WarnBg = '#FDE7E9'; $card.WarnFg = '#B10E1C'; $card.WarnVis = 'Visible' }
                elseif ($m -gt 0) { $card.Warn = (T 'warn_merge') -f $m; $card.WarnBg = '#FFE8D6'; $card.WarnFg = '#8A3707'; $card.WarnVis = 'Visible' }
            }
            'create' {
                $card.From = $it.Father; $card.To = $it.Code; $card.Name = $it.Name; $card.Stats = (T 'card_create') -f $it.Father
                $card.Rule = (T 'rule_t_create') + ' · ' + $it.Status; $card.RuleBg = '#DFF6DD'; $card.RuleFg = '#0E700E'
                if ($it.Kind -eq 'invalid') { $card.Warn = (T 'warn_conf') -f 1; $card.WarnBg = '#FDE7E9'; $card.WarnFg = '#B10E1C'; $card.WarnVis = 'Visible' }
            }
            'rename' {
                $card.From = $it.Code; $card.To = ''; $card.ArrowVis = 'Collapsed'; $card.Name = $it.NewName; $card.Stats = (T 'card_rename') -f $it.OldName
                $card.Rule = T 'rule_t_rename'; $card.RuleBg = '#F1E9FA'; $card.RuleFg = '#5C2E91'
            }
        }
        $cards.Add($card)
    }
    ,$cards
}
function Get-PlanTotals {
    $tas = 0; $bad = 0; $merges = 0
    foreach ($it in $script:Plan) {
        if ($it.Type -eq 'move') { $tas += @($it.Pairs).Count; $bad += @($it.Pairs | Where-Object { $_.Kind -eq 'invalid' }).Count; $merges += @($it.Pairs | Where-Object { $_.Kind -eq 'merge' }).Count }
        elseif ($it.Type -eq 'create') { $tas++; if ($it.Kind -eq 'invalid') { $bad++ } }
    }
    @{ TAs = $tas; Bad = $bad; Merges = $merges }
}
function New-ExecConfirmArgs {
    $tot = Get-PlanTotals
    $stats = New-Object 'System.Collections.Generic.List[MoveTA.Metric]'
    foreach ($s in @(@((T 'cf_tas'), [string]$tot.TAs, '#0F548C'), @((T 'cf_moves'), [string]$script:Plan.Count, '#242424'), @((T 'cf_plant'), $script:PlantId, '#242424'))) {
        $x = New-Object MoveTA.Metric; $x.Label = $s[0]; $x.Value = $s[1]; $x.Fg = $s[2]; $stats.Add($x)
    }
    $more = $(if ($script:Plan.Count -gt 6) { (T 'cf_more') -f ($script:Plan.Count - 6) } else { '' })
    @{ Title = (T 'cf_exec_title'); Sub = ((T 'cf_exec_sub') + "`n" + (T 'cf_db') + ': ' + $script:ConnLabel); OkText = (T 'cf_go'); Kind = 'warn'
       Stats = $stats; Moves = (Get-PlanCards 6); More = $more; Warn = (T 'cf_backup'); OkGlyph = [string][char]0xE73E }
}
function Set-Busy([bool]$On) {
    [System.Windows.Input.Mouse]::OverrideCursor = $(if ($On) { [System.Windows.Input.Cursors]::Wait } else { $null })
    if ($On) { $script:Win.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render) }
}
function Switch-Panel([string]$Name) {
    foreach ($r in @($script:U.RailMove, $script:U.RailHist, $script:U.RailConn)) {
        if ($r.Tag -eq $Name) { $r.Background = B '#FFFFFF'; $r.Foreground = B '#0F6CBD' } else { $r.Background = B 'Transparent'; $r.Foreground = B '#424242' }
    }
    $script:U.PanelMove.Visibility = $(if ($Name -eq 'move') { 'Visible' } else { 'Collapsed' })
    $script:U.PanelHist.Visibility = $(if ($Name -eq 'hist') { 'Visible' } else { 'Collapsed' })
    $script:U.PanelConn.Visibility = $(if ($Name -eq 'conn') { 'Visible' } else { 'Collapsed' })
    if ($Name -eq 'hist') { Load-History }
}
function Save-Visual($Visual, [string]$Path) {
    $Visual.UpdateLayout()
    $w = [int][Math]::Ceiling($Visual.ActualWidth); $h = [int][Math]::Ceiling($Visual.ActualHeight)
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($Visual)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $fs = [IO.File]::Create($Path); try { $enc.Save($fs) } finally { $fs.Close() }
    Write-Host "[ui] print: $Path"
}
function Get-NodeFromSource($o) {
    $guard = 0
    while ($o -and $guard -lt 60) {
        if ($o -is [System.Windows.FrameworkElement] -and $o.DataContext -is [MoveTA.TaNode]) { return $o.DataContext }
        if ($o -is [System.Windows.Media.Visual]) { $o = [System.Windows.Media.VisualTreeHelper]::GetParent($o) }
        elseif ($o -is [System.Windows.FrameworkContentElement]) { $o = $o.Parent }
        else { $o = $null }
        $guard++
    }
    $null
}
function Get-Descendant($El, [type]$T) {
    if (-not $El) { return $null }
    $n = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($El)
    for ($i = 0; $i -lt $n; $i++) { $c = [System.Windows.Media.VisualTreeHelper]::GetChild($El, $i); if ($c -is $T) { return $c }; $r = Get-Descendant $c $T; if ($r) { return $r } }
    $null
}
function Show-NodeInTree($Tree, $Node) {
    if (-not $Node) { return }
    $path = New-Object System.Collections.ArrayList; $n = $Node
    while ($n) { [void]$path.Insert(0, $n); $n = $n.Parent }
    foreach ($p in $path) { if ($p -ne $Node) { $p.IsExpanded = $true } }
    $container = $Tree; $item = $null
    foreach ($p in $path) {
        $container.UpdateLayout()
        $item = $container.ItemContainerGenerator.ContainerFromItem($p)
        if (-not $item) { break }
        $container = $item
    }
    if ($item) { $item.BringIntoView(); $Node.IsSelected = $true }
}

# ============================================================================ arvores
function Add-Ghost($Idx, $Roots, [string]$Code, [string]$ParentCode, [string]$Name, [string]$Links, [string]$Kind, [string]$Chip) {
    $parent = $Idx[$ParentCode]
    $g = New-Object MoveTA.TaNode; $g.Code = $Code; $g.Name = $Name; $g.Links = $Links; $g.Father = $ParentCode
    $g.State = $Kind; $g.ChipKind = $Kind; $g.ChipText = $Chip
    if ($parent) { $coll = $parent.Children } else { $coll = $Roots }
    $pos = $coll.Count; for ($i = 0; $i -lt $coll.Count; $i++) { if ([string]::CompareOrdinal($coll[$i].Code, $g.Code) -gt 0) { $pos = $i; break } }
    $coll.Insert($pos, $g); $g.Parent = $parent; $Idx[$Code] = $g
    $g
}
function Build-Tree([bool]$WithPlan) {
    $idx = @{}
    $roots = New-Object 'System.Collections.ObjectModel.ObservableCollection[MoveTA.TaNode]'
    foreach ($code in $script:SortedCodes) {
        $b = $script:Base[$code]
        $n = New-Object MoveTA.TaNode; $n.Code = $code; $n.Name = $b.Name; $n.Father = $b.Father; $n.Links = $b.LinkText
        $idx[$code] = $n
    }
    foreach ($code in $script:SortedCodes) {
        $n = $idx[$code]; $p = $null; if ($n.Father) { $p = $idx[$n.Father] }
        if ($p) { $n.Parent = $p; $p.Children.Add($n) } else { $roots.Add($n) }
    }
    if ($WithPlan) {
        foreach ($it in $script:Plan) {
            switch ($it.Type) {
                'move' {
                    $toByFrom = @{}; foreach ($p in $it.Pairs) { $toByFrom[$p.From] = $p.To }
                    foreach ($p in ($it.Pairs | Sort-Object @{ e = { $_.To.Length } }, @{ e = { $_.To } })) {
                        $srcNode = $idx[$p.From]
                        if ($srcNode -and -not $srcNode.IsGhost) { $srcNode.State = 'moved'; $srcNode.ChipKind = 'moved'; $srcNode.ChipText = '→ ' + $p.To }
                        if (-not $p.To -or $p.Kind -eq 'invalid') { continue }
                        $nm = $(if ($p.ToName) { $p.ToName } else { $p.Name })
                        if ($p.Kind -eq 'merge' -and $idx.ContainsKey($p.To)) {
                            $t = $idx[$p.To]; if (-not $t.IsGhost) { $t.State = 'merge' }
                            $t.ChipKind = 'merge'; $t.ChipText = 'MERGE ← ' + $p.From
                            continue
                        }
                        if ($p.From -eq $it.Src) { $parentCode = $it.Target } else { $parentCode = $toByFrom[$script:Base[$p.From].Father] }
                        [void](Add-Ghost $idx $roots $p.To $parentCode $nm $p.Links $p.Kind $p.Status)
                    }
                    $n = $idx[$it.Target]; while ($n) { $n.IsExpanded = $true; $n = $n.Parent }
                    if ($idx.ContainsKey($it.NewTop)) { $idx[$it.NewTop].IsExpanded = $true }
                }
                'create' {
                    if ($it.Kind -ne 'invalid' -and -not $idx.ContainsKey($it.Code)) {
                        $g = Add-Ghost $idx $roots $it.Code $it.Father $it.Name '' $it.Kind $it.Status
                        $n = $g.Parent; while ($n) { $n.IsExpanded = $true; $n = $n.Parent }
                    }
                }
                'rename' {
                    if ($idx.ContainsKey($it.Code)) {
                        $t = $idx[$it.Code]; $t.Name = $it.NewName
                        if ($t.State -eq 'normal') { $t.State = 'renamed' }
                        if (-not $t.ChipText) { $t.ChipKind = 'rename'; $t.ChipText = T 'st_rename' }
                    }
                }
            }
        }
    }
    @{ Roots = $roots; Index = $idx }
}
function Build-SrcTree {
    $open = @{}; foreach ($n in $script:SrcIndex.Values) { if ($n.IsExpanded) { $open[$n.Code] = 1 } }
    $r = Build-Tree $false
    $script:SrcIndex = $r.Index
    foreach ($c in $open.Keys) { if ($r.Index.ContainsKey($c)) { $r.Index[$c].IsExpanded = $true } }
    foreach ($n in $r.Roots) { $n.IsExpanded = $true }
    $script:U.TreeSrc.ItemsSource = $r.Roots
    $script:U.TxtSrcCount.Text = (T 'src_count') -f $script:Base.Count
}
function Build-DstTree {
    $open = @{}; foreach ($n in $script:DstIndex.Values) { if ($n.IsExpanded) { $open[$n.Code] = 1 } }
    $r = Build-Tree $true
    foreach ($c in $open.Keys) { if ($r.Index.ContainsKey($c)) { $r.Index[$c].IsExpanded = $true } }
    foreach ($n in $r.Roots) { $n.IsExpanded = $true }
    $script:DstIndex = $r.Index
    $script:U.TreeDst.ItemsSource = $r.Roots
}
function Update-CreateItem($It, [int]$Upto) {
    $planned = Get-PlannedTargets $Upto
    if ($script:Base.ContainsKey($It.Code) -or $planned.ContainsKey($It.Code)) { $It.Kind = 'invalid'; $It.Status = T 'st_conflict' }
    elseif ($script:ClassSet.ContainsKey($It.Code)) { $It.Kind = 'activate'; $It.Status = T 'st_activate' }
    elseif ($script:DeletedSet.ContainsKey($It.Code)) { $It.Kind = 'reactivate'; $It.Status = T 'st_react' }
    else { $It.Kind = 'new'; $It.Status = T 'st_new' }
}
function Refresh-Plan {
    for ($i = 0; $i -lt $script:Plan.Count; $i++) {
        $it = $script:Plan[$i]
        switch ($it.Type) {
            'move' {
                $it.Pairs = Get-Mapping $it.Src $it.NewTop $i; $it.RuleText = T ('rule_t_' + $it.Rule)
                foreach ($p in $it.Pairs) { if ($p.To -and $it.NameOverrides.ContainsKey($p.To)) { $p.ToName = $it.NameOverrides[$p.To] } }
            }
            'create' { Update-CreateItem $it $i }
        }
    }
    $script:U.PlanList.ItemsSource = Get-PlanCards
    $script:U.PlanEmpty.Visibility = $(if ($script:Plan.Count) { 'Collapsed' } else { 'Visible' })
    $script:U.TxtPlanCount.Text = $(if ($script:Plan.Count) { (T 'plan_count') -f $script:Plan.Count } else { '' })
    $sources = Get-PlannedSources
    foreach ($n in $script:SrcIndex.Values) {
        if ($sources.ContainsKey($n.Code)) { $n.State = 'moved'; $n.ChipKind = 'moved'; $n.ChipText = '→ ' + $sources[$n.Code] }
        elseif ($n.State -ne 'normal') { $n.State = 'normal'; $n.ChipKind = $null; $n.ChipText = $null }
    }
    Build-DstTree
    $tot = Get-PlanTotals
    if ($script:Plan.Count) {
        $s = (T 'plan_sum') -f $script:Plan.Count, $tot.TAs
        if ($tot.Merges) { $s += (T 'sum_merge') -f $tot.Merges }
        if ($tot.Bad) { $s += (T 'sum_bad') -f $tot.Bad }
        $script:U.TxtPlanSummary.Text = $s
    } else { $script:U.TxtPlanSummary.Text = T 'plan_emptysum' }
    $script:U.BtnSim.IsEnabled = ($script:Plan.Count -gt 0 -and $tot.Bad -eq 0 -and $script:Conn)
    $script:U.BtnExec.IsEnabled = ($script:SimOk -and $script:Plan.Count -gt 0 -and $tot.Bad -eq 0)
    $script:U.BtnUndo.IsEnabled = ($script:Plan.Count -gt 0)
    $script:U.BtnClear.IsEnabled = ($script:Plan.Count -gt 0)
}
function Clear-Search { foreach ($ix in @($script:SrcIndex, $script:DstIndex)) { foreach ($n in $ix.Values) { if ($n.IsMatch) { $n.IsMatch = $false } } } }
function Apply-Search([string]$Q) {
    Clear-Search
    $Q = $Q.Trim(); if (-not $Q) { return }
    $first = $null; $count = 0
    foreach ($ix in @($script:SrcIndex, $script:DstIndex)) {
        foreach ($code in ($ix.Keys | Sort-Object)) {
            $n = $ix[$code]
            if ($n.Code.StartsWith($Q, [StringComparison]::OrdinalIgnoreCase) -or ($n.Name -and $n.Name.IndexOf($Q, [StringComparison]::OrdinalIgnoreCase) -ge 0)) {
                $n.IsMatch = $true; $p = $n.Parent; while ($p) { $p.IsExpanded = $true; $p = $p.Parent }
                if ($ix -eq $script:SrcIndex) { $count++; if (-not $first) { $first = $n } }
            }
        }
    }
    if ($first) { Show-NodeInTree $script:U.TreeSrc $first; Show-Info 'info' (T 'search_t') ((T 'search_found') -f $count, $Q) }
    else { Show-Info 'warn' (T 'search_t') ((T 'search_none') -f $Q) }
}

# ============================================================================ dialogo mover
function New-MoveDialog([string]$Src, [string]$Tgt) {
    $w = New-Window $script:DialogXaml
    $c = @{}
    foreach ($n in @('DlgRoot','DlgHeader','TxtSub','TxtSrcCode','TxtSrcName','TxtStats','TxtTgtCode','TxtTgtName','ChipTgtNew','RbPrefix','RbNext','RbCustom','TxtPrefixDesc','TxtNextDesc',
                     'TbCustom','TxtCustomDesc','TxtPrevHdr','ChipNew','CntNew','ChipReact','CntReact','ChipMerge','CntMerge','ChipBad','CntBad','GridPrev','TxtErr','BtnOk','BtnCancel','BtnX','TxtFoot')) { $c[$n] = $w.FindName($n) }
    Set-ColumnHeaders $c.GridPrev @('col_cur','col_newc','col_name','col_links','col_status')
    if (-not $UiTest -and $script:Win.IsLoaded) { $w.Owner = $script:Win }
    $suf = Get-Suffix $Src
    $sep = [regex]::Match($suf, '^[^0-9A-Za-z]*').Value
    $width = $suf.Length - $sep.Length
    $ctx = [pscustomobject]@{ Window = $w; C = $c; Src = $Src; Tgt = $Tgt; Suffix = $suf; Sep = $sep; Width = $width
                              PrefixTop = ($Tgt + $suf); NextTop = (Get-NextFree $Tgt $sep $width); NewTop = $null; Rule = $null; Mapping = $null }
    $script:DlgCtx = $ctx
    $c.TxtSrcCode.Text = $Src; $c.TxtSrcName.Text = $script:Base[$Src].Name
    $c.TxtTgtCode.Text = $Tgt; $c.TxtTgtName.Text = Get-NameOf $Tgt
    if ($script:DstIndex.ContainsKey($Tgt) -and $script:DstIndex[$Tgt].IsGhost) { $c.ChipTgtNew.Visibility = 'Visible' }
    $sub = Get-Subtree $Src
    $c.TxtStats.Text = Get-StatsText (Get-Mapping $Src $Src)
    $c.TxtSub.Text = (T 'dlg_sub') -f $script:PlantId, $sub.Count
    $kid = $null; if ($sub.Count -gt 1) { $kid = $sub[1] }
    $ex = { param($top) if ($kid) { "   ·   $kid → $($top + $kid.Substring($Src.Length))" } else { '' } }
    $c.TxtPrefixDesc.Text = "$Src → $($ctx.PrefixTop)" + (& $ex $ctx.PrefixTop)
    if ($ctx.NextTop) { $c.TxtNextDesc.Text = ((T 'desc_next') -f $Src, $ctx.NextTop, $Tgt) + (& $ex $ctx.NextTop) }
    else { $c.TxtNextDesc.Text = T 'desc_next_na'; $c.RbNext.IsEnabled = $false }
    $c.TbCustom.Text = $Tgt + $sep
    $c.TxtCustomDesc.Text = (T 'desc_custom') -f ($Tgt + $sep + ('N' * [Math]::Max($width, 1)))
    $c.RbPrefix.Add_Checked({ Update-Dialog }); $c.RbNext.Add_Checked({ Update-Dialog })
    $c.RbCustom.Add_Checked({ Update-Dialog; $tb = $script:DlgCtx.C.TbCustom; $tb.Focus() | Out-Null; $tb.CaretIndex = $tb.Text.Length })
    $c.TbCustom.Add_TextChanged({ if ($script:DlgCtx.C.RbCustom.IsChecked) { Update-Dialog } })
    $c.TbCustom.Add_GotKeyboardFocus({ if (-not $script:DlgCtx.C.RbCustom.IsChecked) { $script:DlgCtx.C.RbCustom.IsChecked = $true } })
    $c.BtnOk.Add_Click({ $script:DlgCtx.Window.DialogResult = $true })
    $c.BtnCancel.Add_Click({ $script:DlgCtx.Window.DialogResult = $false })
    $c.BtnX.Add_Click({ $script:DlgCtx.Window.DialogResult = $false })
    $c.DlgHeader.Add_MouseLeftButtonDown({ try { $script:DlgCtx.Window.DragMove() } catch {} })
    $w.Add_KeyDown({ param($s, $e)
        if ($e.Key -eq 'Escape') { $script:DlgCtx.Window.DialogResult = $false }
        elseif ($e.Key -eq 'Return' -and $script:DlgCtx.C.BtnOk.IsEnabled) { $script:DlgCtx.Window.DialogResult = $true }
    })
    if ($ctx.PrefixTop -eq $Src) { if ($ctx.NextTop) { $c.RbNext.IsChecked = $true } else { $c.RbCustom.IsChecked = $true } }
    else { $c.RbPrefix.IsChecked = $true }
    Update-Dialog
    $ctx
}
function Update-Dialog {
    $d = $script:DlgCtx; $c = $d.C
    $rule = $(if ($c.RbPrefix.IsChecked) { 'prefix' } elseif ($c.RbNext.IsChecked) { 'next' } else { 'custom' })
    $top = $null; $err = $null
    switch ($rule) {
        'prefix' { $top = $d.PrefixTop; if ($top -eq $d.Src) { $err = (T 'err_prefix_same') -f $d.Tgt } }
        'next'   { $top = $d.NextTop; if (-not $top) { $err = T 'err_next' } }
        'custom' {
            $top = $c.TbCustom.Text.Trim()
            if (-not $top) { $err = T 'err_empty' }
            elseif (-not $top.StartsWith($d.Tgt)) { $err = (T 'err_start') -f $d.Tgt }
            elseif ($top.Length -ne $d.Tgt.Length + $d.Suffix.Length) { $err = (T 'err_format') -f ($d.Tgt + $d.Sep + ('N' * $d.Width)), $d.Src }
            elseif ($d.Sep -and -not $top.Substring($d.Tgt.Length).StartsWith($d.Sep)) { $err = (T 'err_sep') -f $d.Tgt, $d.Sep }
        }
    }
    if (-not $err -and $top -eq $d.Src) { $err = T 'err_equal' }
    if (-not $err -and $top.StartsWith($d.Src) -and $top.Length -gt $d.Src.Length) { $err = T 'err_inside' }
    $map = $null; if ($top -and -not $err) { $map = Get-Mapping $d.Src $top }
    $rows = New-Object 'System.Collections.Generic.List[MoveTA.PreviewRow]'
    $cnt = @{ new = 0; react = 0; merge = 0; bad = 0 }
    if ($map) {
        foreach ($m in $map) {
            $col = Get-KindColors $m.Kind
            $r = New-Object MoveTA.PreviewRow; $r.From = $m.From; $r.To = $m.To; $r.Name = $m.Name; $r.Links = $m.Links; $r.Status = $m.Status; $r.StatusBg = $col[0]; $r.StatusFg = $col[1]
            $rows.Add($r)
            switch ($m.Kind) { 'new' { $cnt.new++ } 'activate' { $cnt.react++ } 'reactivate' { $cnt.react++ } 'merge' { $cnt.merge++ } default { $cnt.bad++ } }
        }
        if ($cnt.bad -gt 0) { $err = (T 'err_conflict') -f $cnt.bad }
    }
    $c.GridPrev.ItemsSource = $rows
    $c.TxtPrevHdr.Text = $(if ($map) { (T 'prev_hdr') -f $rows.Count } else { T 'prev' })
    $c.CntNew.Text = (T 'cnt_new') -f $cnt.new;       $c.ChipNew.Visibility   = $(if ($cnt.new)   { 'Visible' } else { 'Collapsed' })
    $c.CntReact.Text = (T 'cnt_react') -f $cnt.react; $c.ChipReact.Visibility = $(if ($cnt.react) { 'Visible' } else { 'Collapsed' })
    $c.CntMerge.Text = (T 'cnt_merge') -f $cnt.merge; $c.ChipMerge.Visibility = $(if ($cnt.merge) { 'Visible' } else { 'Collapsed' })
    $c.CntBad.Text = (T 'cnt_bad') -f $cnt.bad;       $c.ChipBad.Visibility   = $(if ($cnt.bad)   { 'Visible' } else { 'Collapsed' })
    if (-not $err -and $cnt.merge -gt 0) { $c.TxtFoot.Text = (T 'foot_merge') -f $cnt.merge } else { $c.TxtFoot.Text = T 'foot' }
    $c.TxtErr.Text = $(if ($err) { $err } else { '' })
    $c.BtnOk.IsEnabled = (-not $err -and $map -and $rows.Count -gt 0)
    $d.NewTop = $top; $d.Rule = $rule; $d.Mapping = $map
}
function Add-PlanItem($Ctx) {
    [void]$script:Plan.Add([pscustomobject]@{ Type = 'move'; Src = $Ctx.Src; Target = $Ctx.Tgt; Rule = $Ctx.Rule; RuleText = (T ('rule_t_' + $Ctx.Rule)); NewTop = $Ctx.NewTop
                                              Pairs = $Ctx.Mapping; Name = $script:Base[$Ctx.Src].Name; NameOverrides = @{} })
    $script:SimOk = $false
    Refresh-Plan
    Show-Info 'info' (T 'info_added_t') ((T 'info_added') -f $Ctx.Src, $Ctx.NewTop, @($Ctx.Mapping).Count)
}
function Open-MoveDialog([string]$Src, [string]$Tgt) {
    $ctx = New-MoveDialog $Src $Tgt
    if ($ctx.Window.ShowDialog() -and $ctx.Mapping) { Add-PlanItem $ctx }
}

# ============================================================================ dialogo nova TA / renomear
function New-EditDialog([string]$Mode, $Node) {
    $w = New-Window $script:EditXaml
    $c = @{}; foreach ($n in @('EdRoot','EdHeader','EdX','EdIconBg','EdIcon','EdTitle','EdSub','EdCtxLabel','EdCtxCode','EdCtxName','EdCode','EdName','EdHint','EdStatus','EdChip','EdChipText','EdStatusText','EdErr','EdCancel','EdOk','EdOkText')) { $c[$n] = $w.FindName($n) }
    $ctx = [pscustomobject]@{ Window = $w; C = $c; Mode = $Mode; Code = $Node.Code; IsGhost = $Node.IsGhost; OldName = $Node.Name; OtherPlants = 0
                              CodeMode = 'locked'; Parent = $null; FixedLen = 0; Fmt = ''; Owner = $null; Count = 1; LockMsg = '' }
    $script:EdCtx = $ctx
    $c.EdSub.Text = (T 'ed_sub') -f $script:PlantId
    if ($Mode -eq 'create') {
        $c.EdTitle.Text = T 'ed_new_t'; $c.EdIconBg.Background = B '#DFF6DD'; $c.EdIcon.Foreground = B '#0E700E'; $c.EdIcon.Text = [string][char]0xE710
        $c.EdCtxLabel.Text = T 'ed_inside'; $c.EdCtxCode.Text = $Node.Code; $c.EdCtxName.Text = $Node.Name
        $pat = Get-ChildPattern $Node.Code
        $sugg = Get-NextFree $Node.Code $pat[0] $pat[1]
        $c.EdCode.Text = $(if ($sugg) { $sugg } else { $Node.Code + $pat[0] })
        $c.EdHint.Text = (T 'ed_code_hint') -f $Node.Code
        $c.EdOkText.Text = T 'btn_add'
    } else {
        $c.EdTitle.Text = T 'ed_ren_t'; $c.EdIconBg.Background = B '#F1E9FA'; $c.EdIcon.Foreground = B '#5C2E91'; $c.EdIcon.Text = [string][char]0xE8AC
        $c.EdCtxLabel.Text = T 'ed_ta'; $c.EdCtxCode.Text = $Node.Code; $c.EdCtxName.Text = $Node.Name
        $c.EdCode.Text = $Node.Code; $c.EdName.Text = $Node.Name
        # o que define as regras do novo código
        if (-not $Node.IsGhost) {
            $why = Test-DragSource $Node.Code $Node.Code
            if ($why) { $ctx.LockMsg = $why }
            else {
                $ctx.CodeMode = 'renumber'; $ctx.Parent = $script:Base[$Node.Code].Father; $suf = Get-Suffix $Node.Code
                $ctx.FixedLen = $ctx.Parent.Length + $suf.Length; $ctx.Fmt = $ctx.Parent + ($suf -replace '[0-9A-Za-z]', 'N'); $ctx.Count = (Get-Subtree $Node.Code).Count
            }
        } else {
            $own = Find-PlanOwner $Node.Code; $ctx.Owner = $own
            if ($own -and $own.Role -eq 'create') { $ctx.CodeMode = 'create'; $ctx.Parent = $own.Item.Father; $ctx.Fmt = $ctx.Parent + '…'; $ctx.Count = Get-GhostCount $Node.Code }
            elseif ($own -and $own.Role -eq 'movetop') {
                $ctx.CodeMode = 'movetop'; $ctx.Parent = $own.Item.Target; $suf = Get-Suffix $own.Item.Src
                $ctx.FixedLen = $ctx.Parent.Length + $suf.Length; $ctx.Fmt = $ctx.Parent + ($suf -replace '[0-9A-Za-z]', 'N'); $ctx.Count = Get-GhostCount $Node.Code
            }
            elseif ($own) { $ctx.LockMsg = (T 'ed_code_lock_child') -f $own.Item.NewTop }
        }
        $c.EdCode.IsEnabled = ($ctx.CodeMode -ne 'locked')
        $codeHint = $(if ($ctx.CodeMode -eq 'locked') { $ctx.LockMsg } else { (T 'ed_code_hint_ren') -f $ctx.Fmt })
        if ($Node.IsGhost) { $c.EdHint.Text = (T 'ed_ren_ghost') + ' ' + $codeHint }
        else {
            try { $ctx.OtherPlants = [int](Invoke-SqlScalar "SELECT COUNT(DISTINCT Plant_Id) FROM ST_PlantFunction WHERE FuncClass_Id = @fc AND Function_Id = @c AND Rec_Deleted = 0 AND Plant_Id <> @p" @{ '@fc' = $script:FuncClass; '@c' = $Node.Code; '@p' = $script:PlantId }) } catch {}
            $c.EdHint.Text = $(if ($ctx.OtherPlants -gt 0) { (T 'ed_ren_info') -f $script:FuncClass, $ctx.OtherPlants } else { (T 'ed_ren_info0') -f $script:FuncClass }) + ' ' + $codeHint
        }
        $c.EdOkText.Text = $(if ($Node.IsGhost) { T 'btn_apply' } else { T 'btn_add' })
    }
    $c.EdCode.Add_TextChanged({ Update-Edit }); $c.EdName.Add_TextChanged({ Update-Edit })
    $c.EdOk.Add_Click({ $script:EdCtx.Window.DialogResult = $true })
    $c.EdCancel.Add_Click({ $script:EdCtx.Window.DialogResult = $false })
    $c.EdX.Add_Click({ $script:EdCtx.Window.DialogResult = $false })
    $c.EdHeader.Add_MouseLeftButtonDown({ try { $script:EdCtx.Window.DragMove() } catch {} })
    $w.Add_KeyDown({ param($s, $e)
        if ($e.Key -eq 'Escape') { $script:EdCtx.Window.DialogResult = $false }
        elseif ($e.Key -eq 'Return' -and $script:EdCtx.C.EdOk.IsEnabled) { $script:EdCtx.Window.DialogResult = $true }
    })
    Update-Edit
    $ctx
}
function Update-Edit {
    $d = $script:EdCtx; $c = $d.C
    $code = $c.EdCode.Text.Trim(); $name = $c.EdName.Text.Trim()
    $err = $null; $kind = $null; $msg = ''
    if ($d.Mode -eq 'create') {
        $planned = Get-PlannedTargets
        if (-not $code) { $err = T 'ed_err_code' }
        elseif (-not $code.StartsWith($d.Code) -or $code.Length -le $d.Code.Length) { $err = (T 'ed_err_start') -f $d.Code }
        elseif ($script:Base.ContainsKey($code)) { $err = (T 'ed_err_exists') -f $code }
        elseif ($planned.ContainsKey($code)) { $err = (T 'ed_err_planned') -f $code }
        elseif (-not $name -and -not $script:ClassSet.ContainsKey($code)) { $err = T 'ed_err_name' }
        if (-not $err) {
            if ($script:ClassSet.ContainsKey($code)) { $kind = 'activate'; $msg = (T 'ed_st_act') -f $script:FuncClass }
            elseif ($script:DeletedSet.ContainsKey($code)) { $kind = 'reactivate'; $msg = T 'ed_st_react' }
            else { $kind = 'new'; $msg = (T 'ed_st_new') -f $script:FuncClass, $script:PlantId }
        }
    } else {
        $codeChanged = ($code -ne $d.Code); $nameChanged = ($name -cne $d.OldName)
        if (-not $name) { $err = T 'ed_err_name' }
        elseif (-not $codeChanged -and -not $nameChanged) { $err = T 'ed_err_same_all' }
        elseif ($codeChanged) {
            if ($d.CodeMode -eq 'locked') { $err = $d.LockMsg }
            elseif (-not $code -or -not $code.StartsWith($d.Parent) -or $code.Length -le $d.Parent.Length) { $err = (T 'ed_err_start') -f $d.Parent }
            elseif ($d.FixedLen -and $code.Length -ne $d.FixedLen) { $err = (T 'ed_err_fmt') -f $d.Fmt }
            elseif ($script:Base.ContainsKey($code)) { $err = (T 'ed_err_exists') -f $code }
            else {
                $pl = Get-PlannedTargets
                if ($d.Owner) { foreach ($k in (Get-CreatedCodes $d.Owner.Item).Keys) { $pl.Remove($k) }; if ($d.Owner.Item.Type -eq 'move') { foreach ($p in $d.Owner.Item.Pairs) { if ($p.To) { $pl.Remove($p.To) } } } }
                if ($pl.ContainsKey($code)) { $err = (T 'ed_err_planned') -f $code }
            }
        }
        if (-not $err) {
            if ($codeChanged) {
                $kind = 'recode'; $msg = (T 'ed_st_recode') -f $d.Code, $code, $d.Count
                if ($nameChanged) { $msg += '  ·  "' + $d.OldName + '" → "' + $name + '"' }
            } else { $kind = 'rename'; $msg = '"' + $d.OldName + '" → "' + $name + '"' }
        }
    }
    if ($name.Length -gt 100) { $err = T 'ed_err_len'; $kind = $null }
    if ($kind) {
        switch ($kind) {
            'new'        { $col = @('#0F6CBD','#FFFFFF','#EBF3FC','#B4D6FA', (T 'st_new')) }
            'activate'   { $col = @('#DFF6DD','#0E700E','#F1FAF1','#9FD89F', (T 'st_activate')) }
            'reactivate' { $col = @('#DFF6DD','#0E700E','#F1FAF1','#9FD89F', (T 'st_react')) }
            'recode'     { $col = @('#0F6CBD','#FFFFFF','#EBF3FC','#B4D6FA', (T 'st_recode')) }
            default      { $col = @('#F1E9FA','#5C2E91','#FAF7FD','#D9C6EE', (T 'st_rename')) }
        }
        $c.EdChip.Background = B $col[0]; $c.EdChipText.Foreground = B $col[1]; $c.EdStatus.Background = B $col[2]; $c.EdStatus.BorderBrush = B $col[3]; $c.EdChipText.Text = $col[4]
        $c.EdStatusText.Text = $msg; $c.EdStatus.Visibility = 'Visible'
    } else { $c.EdStatus.Visibility = 'Collapsed' }
    $c.EdErr.Text = $(if ($err) { $err } else { '' })
    $c.EdOk.IsEnabled = (-not $err)
}
function Apply-Edit($Ctx) {
    $code = $Ctx.C.EdCode.Text.Trim(); $name = $Ctx.C.EdName.Text.Trim()
    if ($Ctx.Mode -eq 'create') {
        $it = [pscustomobject]@{ Type = 'create'; Code = $code; Father = $Ctx.Code; Name = $name; Kind = $null; Status = $null }
        if (-not $name) { $it.Name = $code }
        [void]$script:Plan.Add($it)
        $script:SimOk = $false; Refresh-Plan
        Show-NodeInTree $script:U.TreeDst $script:DstIndex[$code]
        Show-Info 'info' (T 'info_created_t') ((T 'info_created') -f $code, $it.Name, $Ctx.Code)
        return
    }
    if ($code -ne $Ctx.Code) {
        $nameChanged = ($name -cne $Ctx.OldName)
        switch ($Ctx.CodeMode) {
            'renumber' {
                # TA já salva: vira um movimento dentro do mesmo pai (renumeração), filhos e ligações vão juntos
                $baseName = $script:Base[$Ctx.Code].Name
                $old = @($script:Plan | Where-Object { $_.Type -eq 'rename' -and $_.Code -eq $Ctx.Code })
                foreach ($o in $old) { $script:Plan.Remove($o) }
                $no = @{}; if ($name -cne $baseName) { $no[$code] = $name }
                [void]$script:Plan.Add([pscustomobject]@{ Type = 'move'; Src = $Ctx.Code; Target = $Ctx.Parent; Rule = 'renumber'; RuleText = (T 'rule_t_renumber'); NewTop = $code
                                                          Pairs = (Get-Mapping $Ctx.Code $code); Name = $baseName; NameOverrides = $no })
            }
            'create' {
                $it = $Ctx.Owner.Item; $oldCode = $it.Code
                $it.Code = $code; $it.Name = $name
                Rename-PlanPrefix $oldCode $code $it
            }
            'movetop' {
                $it = $Ctx.Owner.Item; $oldTop = $it.NewTop
                Set-MoveTop $it $code; $it.Rule = 'custom'
                if ($nameChanged) { $it.NameOverrides[$code] = $name }
                Rename-PlanPrefix $oldTop $code $it
            }
        }
        $script:SimOk = $false; Refresh-Plan
        Show-NodeInTree $script:U.TreeDst $script:DstIndex[$code]
        Show-Info 'info' (T 'info_recoded_t') ((T 'info_recoded') -f $Ctx.Code, $code, $Ctx.Count)
        return
    }
    $done = $false
    if ($Ctx.IsGhost) {
        foreach ($it in $script:Plan) {
            if ($it.Type -eq 'create' -and $it.Code -eq $Ctx.Code) { $it.Name = $name; $done = $true; break }
            if ($it.Type -eq 'move' -and @($it.Pairs | Where-Object { $_.To -eq $Ctx.Code }).Count) { $it.NameOverrides[$Ctx.Code] = $name; $done = $true; break }
        }
    }
    if (-not $done) {
        $ex = $script:Plan | Where-Object { $_.Type -eq 'rename' -and $_.Code -eq $Ctx.Code } | Select-Object -First 1
        if ($ex) { $ex.NewName = $name } else { [void]$script:Plan.Add([pscustomobject]@{ Type = 'rename'; Code = $Ctx.Code; OldName = $script:Base[$Ctx.Code].Name; NewName = $name }) }
    }
    $script:SimOk = $false; Refresh-Plan
    Show-Info 'info' (T 'info_renamed_t') ((T 'info_renamed') -f $Ctx.Code, $Ctx.OldName, $name)
}
function Open-EditTA([string]$Mode, $Node) {
    if (-not $Node) { Show-Info 'warn' (T 'sel_none_t') (T 'sel_none'); return }
    if ($script:DstIndex.ContainsKey($Node.Code)) { $Node = $script:DstIndex[$Node.Code] }
    if ($Node.State -eq 'moved') { Show-Info 'warn' (T 'sel_none_t') ((T 'ed_cant_moved') -f $Node.Code); return }
    $ctx = New-EditDialog $Mode $Node
    if ($script:Win.IsLoaded) { $ctx.Window.Owner = $script:Win }
    if ($Mode -eq 'create') { [void]$ctx.C.EdName.Focus() }
    elseif ($ctx.CodeMode -ne 'locked') { [void]$ctx.C.EdCode.Focus(); $ctx.C.EdCode.SelectAll() }
    else { [void]$ctx.C.EdName.Focus(); $ctx.C.EdName.SelectAll() }
    if ($ctx.Window.ShowDialog()) { Apply-Edit $ctx }
}

# ============================================================================ execucao
function Get-Options {
    @{ WH = [bool]$script:U.ChkWH.IsChecked; PO = [bool]$script:U.ChkPO.IsChecked; Mat = [bool]$script:U.ChkMat.IsChecked; Other = [bool]$script:U.ChkOther.IsChecked
       Retire = [bool]$script:U.ChkRetire.IsChecked; RetireClass = [bool]$script:U.ChkRetireClass.IsChecked; Items = [bool]$script:U.ChkItems.IsChecked }
}
function Get-ActionText([string]$Acao) { if ($script:Lang -eq 'en' -and $script:ActionEn.ContainsKey($Acao)) { $script:ActionEn[$Acao] } else { $Acao } }
function New-LogRow([string]$Par, [string]$Tabela, [string]$Acao, $Linhas, [string]$Detalhe) {
    $r = New-Object MoveTA.LogRow; $r.Par = $Par; $r.Tabela = $Tabela; $r.Acao = $Acao; $r.AcaoText = Get-ActionText $Acao; $r.Linhas = [string]$Linhas; $r.Detalhe = $Detalhe
    $r.RowBg = $(switch ($Acao) { { $_ -in 'ERRO','ATENCAO','VERIFICAR' } { '#FDF3F4' } { $_ -in 'criada','reativada','desativada','COMMIT' } { '#F1FAF1' } 'renomeada' { '#FAF7FD' } { $_ -in 'MANTIDA','ROLLBACK' } { '#FFF9F5' } default { '#FFFFFF' } })
    $r
}
function Invoke-Plan([bool]$Commit) {
    $opt = Get-Options
    $ops = New-Object System.Collections.ArrayList; $movePairs = @(); $renames = @()
    foreach ($it in $script:Plan) {
        switch ($it.Type) {
            'move'   { foreach ($p in $it.Pairs) { [void]$ops.Add([pscustomobject]@{ Kind = 'move'; From = $p.From; To = $p.To; ToName = $p.ToName }); $movePairs += $p } }
            'create' { [void]$ops.Add([pscustomobject]@{ Kind = 'create'; From = $it.Father; To = $it.Code; ToName = $it.Name }) }
            'rename' { $renames += $it }
        }
    }
    $ordered = $ops | Sort-Object @{ e = { $_.To.Length } }, @{ e = { $_.To } }
    $rows = New-Object 'System.Collections.Generic.List[MoveTA.LogRow]'
    $err = $null
    $tran = $script:Conn.BeginTransaction('MoveTA')
    try {
        foreach ($o in $ordered) {
            if ($o.Kind -eq 'create') {
                $res = Invoke-Sql "EXEC dbo.usp_IPS_CreateTA @Plant_Id, @Code, @Name, @Father, NULL" @{ '@Plant_Id' = $script:PlantId; '@Code' = $o.To; '@Name' = $o.ToName; '@Father' = $o.From } $tran
                foreach ($r in $res[0].Rows) { $rows.Add((New-LogRow ((T 'log_create') -f $o.To) $r.Tabela $r.Acao $r.Linhas $r.Detalhe)) }
            } else {
                $res = Invoke-Sql "EXEC dbo.usp_IPS_MoveTA @Plant_Id, @From, @To, @ToName, 1, @WH, @PO, @Mat, @Other, 0, 0, @Items, NULL, 0" @{
                    '@Plant_Id' = $script:PlantId; '@From' = $o.From; '@To' = $o.To; '@ToName' = $o.ToName; '@WH' = [int]$opt.WH; '@PO' = [int]$opt.PO; '@Mat' = [int]$opt.Mat; '@Other' = [int]$opt.Other; '@Items' = [int]$opt.Items } $tran
                foreach ($r in $res[0].Rows) { $rows.Add((New-LogRow "$($o.From) → $($o.To)" $r.Tabela $r.Acao $r.Linhas $r.Detalhe)) }
            }
        }
        foreach ($rn in $renames) {
            $res = Invoke-Sql "EXEC dbo.usp_IPS_RenameTA @Plant_Id, @Code, @Name, NULL" @{ '@Plant_Id' = $script:PlantId; '@Code' = $rn.Code; '@Name' = $rn.NewName } $tran
            foreach ($r in $res[0].Rows) { $rows.Add((New-LogRow ((T 'log_rename') -f $rn.Code) $r.Tabela $r.Acao $r.Linhas $r.Detalhe)) }
        }
        if ($opt.Retire) {
            foreach ($p in ($movePairs | Sort-Object @{ e = { -$_.From.Length } }, @{ e = { $_.From } })) {
                $res = Invoke-Sql "EXEC dbo.usp_IPS_RetireTA @Plant_Id, @F, @C, NULL" @{ '@Plant_Id' = $script:PlantId; '@F' = $p.From; '@C' = [int]$opt.RetireClass } $tran
                foreach ($r in $res[0].Rows) { $rows.Add((New-LogRow ((T 'log_retire') -f $p.From) $r.Tabela $r.Acao $r.Linhas $r.Detalhe)) }
            }
        }
        if ($Commit) { $tran.Commit(); $rows.Add((New-LogRow '*' '-' 'COMMIT' '' (T 'log_commit'))) }
        else { $tran.Rollback(); $rows.Add((New-LogRow '*' '-' 'ROLLBACK' '' (T 'log_rollback'))) }
    } catch {
        try { $tran.Rollback() } catch {}
        $err = $_.Exception.Message
        $rows.Add((New-LogRow '*' '-' 'ERRO' '' ((T 'log_err') -f $err)))
    }
    [pscustomobject]@{ Rows = $rows; Error = $err; Pairs = (Get-PlanTotals).TAs }
}
function Get-Metrics($Rows, [int]$Pairs) {
    $sum = { param($tab, [string[]]$acoes) $s = 0; foreach ($r in $Rows) { if ($r.Tabela -eq $tab -and $acoes -contains $r.Acao) { $v = 0; [void][int]::TryParse($r.Linhas, [ref]$v); $s += $v } }; $s }
    $cnt = { param($tab, [string[]]$acoes) @($Rows | Where-Object { $_.Tabela -eq $tab -and $acoes -contains $_.Acao }).Count }
    $alerts = @($Rows | Where-Object { $_.Acao -in 'ERRO','ATENCAO','VERIFICAR','MANTIDA' }).Count
    $list = New-Object 'System.Collections.Generic.List[MoveTA.Metric]'
    foreach ($m in @(
        @('m_tas', $Pairs, '#242424'),
        @('m_created', (& $cnt 'ST_PlantFunction' @('criada','reativada')), '#0F548C'),
        @('m_pfo', (& $sum 'ST_PlantFunctionObject' @('movidas')), '#242424'),
        @('m_wp', (& $sum 'ST_WPWHERE2' @('movidas')), '#242424'),
        @('m_wh', (& $sum 'ST_WHWHERE2' @('movidas')), '#242424'),
        @('m_hist', (& $sum 'ST_WorkHistory' @('atualizadas')), '#242424'),
        @('m_renamed', (& $cnt 'ST_Function' @('renomeada')), '#5C2E91'),
        @('m_retired', (& $cnt 'ST_PlantFunction' @('desativada')), '#0E700E'),
        @('m_alerts', $alerts, $(if ($alerts) { '#B10E1C' } else { '#242424' })))) {
        $x = New-Object MoveTA.Metric; $x.Key = $m[0]; $x.Label = T $m[0]; $x.Value = [string]$m[1]; $x.Fg = $m[2]; $list.Add($x)
    }
    ,$list
}
function Get-Resumo($Metrics) {
    $mv = @{}; foreach ($x in $Metrics) { $mv[$x.Key] = $x.Value }
    (T 'resumo') -f $mv['m_tas'], $mv['m_pfo'], $mv['m_wp'], $mv['m_wh'], $mv['m_hist'], $mv['m_alerts']
}
function Save-Log($Rows, [bool]$Commit) {
    if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }
    $kind = $(if ($Commit) { 'EXECUCAO' } else { 'SIMULACAO' })
    $file = Join-Path $script:LogDir ("{0}_{1}_{2}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $kind, $script:PlantId)
    $Rows | Select-Object Par, Tabela, Acao, Linhas, Detalhe | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $file
}
function Update-ActivityHeader {
    if ($script:LastRows) { $script:U.TxtActivity.Text = (T 'act_line') -f (T $script:LastKind), $script:LastTime, $script:LastRows.Count }
    else { $script:U.TxtActivity.Text = T 'act_none' }
}
function Run-Plan([bool]$Commit) {
    if (-not $script:Plan.Count) { return }
    if ($Commit -and -not $UiTest) {
        $a = New-ExecConfirmArgs
        if (-not (Show-Confirm @a)) { return }
    }
    Set-Busy $true
    try { $res = Invoke-Plan $Commit } finally { Set-Busy $false }
    $script:LastRows = $res.Rows; $script:LastPairs = $res.Pairs
    $script:LastKind = $(if ($Commit) { 'act_exec' } else { 'act_sim' }); $script:LastTime = Get-Date -Format 'HH:mm:ss'
    $script:U.GridLog.ItemsSource = $res.Rows
    $metrics = Get-Metrics $res.Rows $res.Pairs
    $script:U.Metrics.ItemsSource = $metrics
    $script:LastLogFile = Save-Log $res.Rows $Commit
    $script:U.BtnSaveLog.IsEnabled = $true
    Update-ActivityHeader
    $resumo = Get-Resumo $metrics
    if ($res.Error) {
        $script:SimOk = $false
        Show-Info 'error' $(if ($Commit) { T 'fail_exec' } else { T 'fail_sim' }) $res.Error
    } elseif (-not $Commit) {
        $script:SimOk = $true
        Show-Info 'success' (T 'sim_ok_t') ((T 'sim_ok') -f $resumo)
    } else {
        $script:SimOk = $false
        $script:Plan.Clear()
        Load-Plant $script:PlantId; Build-SrcTree; $script:DstIndex = @{}
        Show-Info 'success' (T 'exec_ok_t') ((T 'exec_ok') -f $resumo, [IO.Path]::GetFileName($script:LastLogFile))
    }
    Refresh-Plan
}

# ============================================================================ conexao / historico
function Set-ConnStatus([string]$Key, [object[]]$ArgList = @()) {
    $script:ConnStatusKey = $Key; $script:ConnStatusArgs = $ArgList
    $script:U.TxtConnStatus.Text = $(if ($Key) { (T $Key) -f $ArgList } else { '' })
}
function Connect-Db {
    try {
        if ($script:Conn) { try { $script:Conn.Close() } catch {} }
        $srv = $script:U.TbServer.Text.Trim(); $db = $script:U.TbDb.Text.Trim()
        $cs = New-ConnString $srv $db ([bool]$script:U.RbWin.IsChecked) $script:U.TbUser.Text $script:U.PbPass.Password
        $script:Conn = New-Object System.Data.SqlClient.SqlConnection $cs
        $script:Conn.Open()
        $ver = [string](Invoke-SqlScalar "SELECT System_DBVersion FROM ST_System")
        if (-not (Test-Procs)) {
            $ok = $(if ($UiTest) { $true } else { Show-Confirm -Title (T 'cf_proc_t') -Sub ((T 'proc_missing') -f $db) -OkText (T 'cf_install') -Kind 'question' })
            if (-not $ok) { throw (T 'proc_notinst') }
            Install-Procs
        }
        $script:ConnLabel = "$db · $srv"; $script:DbName = $db; $script:IpsVer = $ver
        Save-Settings @{ server = $srv; database = $db }
        $script:U.TxtConnChip.Text = "$db · IPS $ver"; $script:U.DotConn.Fill = B '#6CCB5F'
        Set-ConnStatus 'conn_ok' @($srv, $db, $ver)
        $plants = (Invoke-Sql "SELECT Plant_Id, Plant_Name FROM ST_Plant WHERE Rec_Deleted = 0 ORDER BY Plant_Name")[0]
        $script:Loading = $true
        $script:U.CbPlant.Items.Clear(); $sel = -1; $i = 0
        foreach ($p in $plants.Rows) {
            $item = New-Object System.Windows.Controls.ComboBoxItem; $item.Content = "$($p.Plant_Id)  ·  $($p.Plant_Name)"; $item.Tag = [string]$p.Plant_Id
            [void]$script:U.CbPlant.Items.Add($item)
            if ($Plant -and $p.Plant_Id -eq $Plant) { $sel = $i } elseif ($sel -lt 0 -and $p.Plant_Id -eq 'FMHUB') { $sel = $i }
            $i++
        }
        if ($sel -lt 0 -and $i -gt 0) { $sel = 0 }
        $script:Loading = $false
        $script:U.CbPlant.SelectedIndex = $sel
        Switch-Panel 'move'
        $true
    } catch {
        $script:Loading = $false
        $script:Conn = $null
        $script:U.TxtConnChip.Text = T 'disconnected'; $script:U.DotConn.Fill = B '#BDBDBD'
        Set-ConnStatus 'conn_fail' @($_.Exception.Message)
        Switch-Panel 'conn'
        $false
    }
}
function Select-Plant {
    $item = $script:U.CbPlant.SelectedItem; if (-not $item) { return }
    Set-Busy $true
    try {
        $script:Plan.Clear(); $script:SimOk = $false; $script:DstIndex = @{}; $script:SrcIndex = @{}
        Load-Plant ([string]$item.Tag); Build-SrcTree; Refresh-Plan
        $script:U.GridLog.ItemsSource = $null; $script:U.Metrics.ItemsSource = $null; $script:LastRows = $null; Update-ActivityHeader
        Show-Info 'info' ((T 'plant_loaded_t') -f $script:PlantId) ((T 'plant_loaded') -f $script:Base.Count, $script:FuncClass)
    } catch { Show-Info 'error' (T 'plant_fail') $_.Exception.Message }
    finally { Set-Busy $false }
}
function Load-History {
    $list = New-Object 'System.Collections.Generic.List[MoveTA.HistRow]'
    if (Test-Path $script:LogDir) {
        foreach ($f in (Get-ChildItem $script:LogDir -Filter *.csv | Sort-Object Name -Descending)) {
            $m = [regex]::Match($f.Name, '^(\d{8})_(\d{6})_(SIMULACAO|EXECUCAO)_(.+)\.csv$')
            $h = New-Object MoveTA.HistRow; $h.File = $f.Name; $h.Path = $f.FullName
            if ($m.Success) {
                $d = [datetime]::ParseExact($m.Groups[1].Value + $m.Groups[2].Value, 'yyyyMMddHHmmss', $null)
                $h.Date = $d.ToString($(if ($script:Lang -eq 'en') { 'yyyy-MM-dd HH:mm:ss' } else { 'dd/MM/yyyy HH:mm:ss' })); $h.Plant = $m.Groups[4].Value
                if ($m.Groups[3].Value -eq 'EXECUCAO') { $h.Kind = T 'kind_exec'; $h.KindBg = '#0F6CBD'; $h.KindFg = '#FFFFFF' } else { $h.Kind = T 'kind_sim'; $h.KindBg = '#EBF3FC'; $h.KindFg = '#0F548C' }
            } else { $h.Date = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'); $h.Kind = 'LOG'; $h.KindBg = '#F0F0F0'; $h.KindFg = '#616161' }
            $list.Add($h)
        }
    }
    $script:U.GridLogs.ItemsSource = $list
    $script:U.TxtHistEmpty.Visibility = $(if ($list.Count) { 'Collapsed' } else { 'Visible' })
}

# ============================================================================ menus de contexto
function New-MenuItem([string]$Key, [int]$Glyph, [string]$Color, [scriptblock]$OnClick) {
    $mi = New-Object System.Windows.Controls.MenuItem
    $mi.Header = T $Key; $mi.Tag = $Key
    $ic = New-Object System.Windows.Controls.TextBlock; $ic.Text = [string][char]$Glyph; $ic.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe Fluent Icons, Segoe MDL2 Assets'); $ic.Foreground = B $Color; $ic.FontSize = 13
    $mi.Icon = $ic; $mi.Add_Click($OnClick)
    $script:MenuItems += $mi
    $mi
}
function Add-TreeMenu($Tree) {
    $cm = New-Object System.Windows.Controls.ContextMenu
    [void]$cm.Items.Add((New-MenuItem 'mn_new' 0xE710 '#0E700E' { Open-EditTA 'create' $script:CtxNode }))
    [void]$cm.Items.Add((New-MenuItem 'mn_rename' 0xE8AC '#5C2E91' { Open-EditTA 'rename' $script:CtxNode }))
    $Tree.ContextMenu = $cm
    $Tree.Add_PreviewMouseRightButtonDown({ param($s, $e)
        $n = Get-NodeFromSource $e.OriginalSource
        $script:CtxNode = $n
        if ($n) { $n.IsSelected = $true }
    })
    $Tree.Add_ContextMenuOpening({ param($s, $e)
        $n = $script:CtxNode
        if (-not $n) { $e.Handled = $true; return }
        $dn = $(if ($script:DstIndex.ContainsKey($n.Code)) { $script:DstIndex[$n.Code] } else { $n })
        foreach ($mi in $s.ContextMenu.Items) { $mi.IsEnabled = ($dn.State -ne 'moved') }
    })
    $Tree.Add_SelectedItemChanged({ param($s, $e) if ($e.NewValue -is [MoveTA.TaNode]) { $script:SelNode = $e.NewValue } })
}

# ============================================================================ idioma
function Apply-Language([string]$L) {
    $script:Lang = $L
    Set-LangResources $script:Win
    Set-ColumnHeaders $script:U.GridLog @('col_move','col_table','col_action','col_rows','col_detail')
    Set-ColumnHeaders $script:U.GridLogs @('col_date','col_kind','col_plant','col_file')
    foreach ($mi in $script:MenuItems) { $mi.Header = T $mi.Tag }
    $script:U.BtnLang.Content = $(if ($L -eq 'pt') { 'EN' } else { 'PT' })
    $script:U.TxtDstHint.Text = T 'dst_hint'
    if (-not $script:Conn) { $script:U.TxtConnChip.Text = T 'disconnected' }
    if ($script:ConnStatusKey) { Set-ConnStatus $script:ConnStatusKey $script:ConnStatusArgs }
    if ($script:Base.Count) { Update-LinkTexts; Build-SrcTree; Refresh-Plan } else { $script:U.TxtPlanSummary.Text = T 'plan_emptysum' }
    if ($script:LastRows) {
        foreach ($r in $script:LastRows) { $r.AcaoText = Get-ActionText $r.Acao }
        $script:U.GridLog.ItemsSource = $null; $script:U.GridLog.ItemsSource = $script:LastRows
        $script:U.Metrics.ItemsSource = Get-Metrics $script:LastRows $script:LastPairs
    }
    Update-ActivityHeader
    if ($script:U.PanelHist.Visibility -eq 'Visible') { Load-History }
    Save-Lang
}

# ============================================================================ eventos
$U.BtnMin.Add_Click({ $script:Win.WindowState = 'Minimized' })
$U.BtnMax.Add_Click({ $script:Win.WindowState = $(if ($script:Win.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }) })
$U.BtnClose.Add_Click({ $script:Win.Close() })
$U.BtnLang.Add_Click({ Apply-Language $(if ($script:Lang -eq 'pt') { 'en' } else { 'pt' }); Show-Info 'info' (T 'lang_t') (T 'lang_set') })
$win.Add_StateChanged({
    $max = $script:Win.WindowState -eq 'Maximized'
    $script:U.Root.Margin = $(if ($max) { New-Object System.Windows.Thickness(7) } else { New-Object System.Windows.Thickness(0) })
    $script:U.BtnMax.Content = $(if ($max) { [string][char]0xE923 } else { [string][char]0xE922 })
})
$win.Add_Closing({ param($s, $e)
    if (-not $UiTest -and $script:Plan.Count -gt 0) {
        if (-not (Show-Confirm -Title (T 'cf_close_t') -Sub (T 'close_confirm') -OkText (T 'cf_close_btn') -Kind 'warn')) { $e.Cancel = $true }
    }
})
$win.Add_Closed({ if ($script:Conn) { try { $script:Conn.Close() } catch {} } })
$win.Add_PreviewKeyDown({ param($s, $e)
    if ($e.OriginalSource -is [System.Windows.Controls.TextBox]) { return }
    if ($e.Key -eq 'Z' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { if ($script:Plan.Count) { Remove-PlanItem ($script:Plan.Count - 1) }; $e.Handled = $true }
    elseif ($e.Key -eq 'F5' -and $script:Conn) { Select-Plant; $e.Handled = $true }
    elseif ($e.Key -eq 'F2' -and $script:SelNode) { Open-EditTA 'rename' $script:SelNode; $e.Handled = $true }
})

$U.RailMove.Add_Click({ Switch-Panel 'move' }); $U.RailHist.Add_Click({ Switch-Panel 'hist' }); $U.RailConn.Add_Click({ Switch-Panel 'conn' }); $U.RailAbout.Add_Click({ Show-About })
$U.BtnInfoClose.Add_Click({ $script:U.InfoBar.Visibility = 'Collapsed' })
$U.RbSql.Add_Checked({ $script:U.TbUser.IsEnabled = $true; $script:U.PbPass.IsEnabled = $true })
$U.RbWin.Add_Checked({ $script:U.TbUser.IsEnabled = $false; $script:U.PbPass.IsEnabled = $false })
$U.BtnConnect.Add_Click({ Set-Busy $true; try { [void](Connect-Db) } finally { Set-Busy $false } })
$U.BtnInstall.Add_Click({
    if (-not $script:Conn) { Set-ConnStatus 'connect_first'; return }
    try { Install-Procs; Set-ConnStatus 'reinstalled' @($script:ConnLabel) } catch { Set-ConnStatus 'conn_fail' @($_.Exception.Message) }
})
$U.CbPlant.Add_SelectionChanged({
    if ($script:Loading) { return }
    if ($script:Plan.Count -gt 0 -and -not $UiTest) {
        if (-not (Show-Confirm -Title (T 'cf_plant_t') -Sub (T 'switch_plant') -OkText (T 'cf_discard') -Kind 'warn')) {
            $script:Loading = $true
            for ($i = 0; $i -lt $script:U.CbPlant.Items.Count; $i++) { if ($script:U.CbPlant.Items[$i].Tag -eq $script:PlantId) { $script:U.CbPlant.SelectedIndex = $i } }
            $script:Loading = $false
            return
        }
    }
    Select-Plant
})
$U.BtnSim.Add_Click({ Run-Plan $false })
$U.BtnExec.Add_Click({ Run-Plan $true })
$U.BtnNewTA.Add_Click({ Open-EditTA 'create' $script:SelNode })
$U.BtnRename.Add_Click({ Open-EditTA 'rename' $script:SelNode })
$U.BtnUndo.Add_Click({ if ($script:Plan.Count) { Remove-PlanItem ($script:Plan.Count - 1) } })
$U.BtnClear.Add_Click({ $script:Plan.Clear(); $script:SimOk = $false; Refresh-Plan; $script:U.InfoBar.Visibility = 'Collapsed' })
$U.BtnReload.Add_Click({ if ($script:Conn) { Select-Plant } })
$U.BtnSaveLog.Add_Click({
    if (-not $script:LastLogFile) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog; $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = [IO.Path]::GetFileName($script:LastLogFile)
    if ($dlg.ShowDialog($script:Win)) { Copy-Item $script:LastLogFile $dlg.FileName -Force; Show-Info 'success' (T 'log_exported') $dlg.FileName }
})
$U.PlanList.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{ param($s, $e)
    $idx = $e.OriginalSource.Tag
    if ($null -ne $idx -and [int]$idx -lt $script:Plan.Count) { Remove-PlanItem ([int]$idx) }
})
$U.TbSearch.Add_TextChanged({ $script:U.TxtSearchHint.Visibility = $(if ($script:U.TbSearch.Text) { 'Collapsed' } else { 'Visible' }) })
$U.TbSearch.Add_KeyDown({ param($s, $e)
    if ($e.Key -eq 'Return') { Apply-Search $script:U.TbSearch.Text }
    elseif ($e.Key -eq 'Escape') { $script:U.TbSearch.Text = ''; Clear-Search; $script:U.InfoBar.Visibility = 'Collapsed' }
})
$U.BtnOpenLog.Add_Click({ $h = $script:U.GridLogs.SelectedItem; if ($h) { Invoke-Item $h.Path } })
$U.GridLogs.Add_MouseDoubleClick({ $h = $script:U.GridLogs.SelectedItem; if ($h) { Invoke-Item $h.Path } })
$U.BtnOpenFolder.Add_Click({ if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }; Invoke-Item $script:LogDir })
$U.BtnRefreshLogs.Add_Click({ Load-History })
Add-TreeMenu $U.TreeSrc
Add-TreeMenu $U.TreeDst

# ---------------- arrastar e soltar / drag and drop
$U.TreeSrc.Add_PreviewMouseLeftButtonDown({ param($s, $e)
    $script:DragStart = $e.GetPosition($script:U.TreeSrc)
    $script:DragNode = Get-NodeFromSource $e.OriginalSource
})
$U.TreeSrc.Add_PreviewMouseMove({ param($s, $e)
    if ($e.LeftButton -ne 'Pressed' -or -not $script:DragNode) { return }
    $p = $e.GetPosition($script:U.TreeSrc)
    if ([Math]::Abs($p.X - $script:DragStart.X) -lt 6 -and [Math]::Abs($p.Y - $script:DragStart.Y) -lt 6) { return }
    $node = $script:DragNode; $script:DragNode = $null
    $why = Test-DragSource $node.Code
    if ($why) { Show-Info 'warn' (T 'drag_cant_t') $why; return }
    $n = (Get-Subtree $node.Code).Count
    $script:U.TxtDstHint.Text = (T 'dragging') -f $node.Code, $n
    $script:U.TxtDstHint.Foreground = B '#0F6CBD'
    $data = New-Object System.Windows.DataObject('MoveTA.Code', $node.Code)
    [void][System.Windows.DragDrop]::DoDragDrop($script:U.TreeSrc, $data, [System.Windows.DragDropEffects]::Move)
    if ($script:DropNode) { $script:DropNode.IsDropTarget = $false; $script:DropNode = $null }
    $script:U.TxtDstHint.Text = T 'dst_hint'; $script:U.TxtDstHint.Foreground = B '#616161'
})
$U.TreeDst.Add_DragOver({ param($s, $e)
    $e.Handled = $true
    $src = [string]$e.Data.GetData('MoveTA.Code')
    $tgt = Get-NodeFromSource $e.OriginalSource
    $why = $(if (-not $src) { T 'drag_from' } else { Test-DropTarget $src $tgt })
    $new = $(if ($why) { $null } else { $tgt })
    if ($script:DropNode -and $script:DropNode -ne $new) { $script:DropNode.IsDropTarget = $false }
    if ($new) { $new.IsDropTarget = $true }
    $script:DropNode = $new
    if ($why) { $e.Effects = [System.Windows.DragDropEffects]::None; $script:U.TxtDstHint.Text = $why }
    else { $e.Effects = [System.Windows.DragDropEffects]::Move; $script:U.TxtDstHint.Text = (T 'drop_here') -f $src, $tgt.Code, $tgt.Name }
    if ($tgt -and $tgt -eq $script:HoverNode) {
        if (-not $tgt.IsExpanded -and $tgt.Children.Count -gt 0 -and ((Get-Date) - $script:HoverSince).TotalMilliseconds -gt 800) { $tgt.IsExpanded = $true }
    } else { $script:HoverNode = $tgt; $script:HoverSince = Get-Date }
    if (-not $script:DstScroll) { $script:DstScroll = Get-Descendant $script:U.TreeDst ([System.Windows.Controls.ScrollViewer]) }
    if ($script:DstScroll) {
        $y = $e.GetPosition($script:U.TreeDst).Y
        if ($y -lt 30) { $script:DstScroll.LineUp() } elseif ($y -gt $script:U.TreeDst.ActualHeight - 30) { $script:DstScroll.LineDown() }
    }
})
$U.TreeDst.Add_DragLeave({ param($s, $e) if ($script:DropNode) { $script:DropNode.IsDropTarget = $false; $script:DropNode = $null } })
$U.TreeDst.Add_Drop({ param($s, $e)
    $e.Handled = $true
    $src = [string]$e.Data.GetData('MoveTA.Code')
    $tgt = Get-NodeFromSource $e.OriginalSource
    if ($script:DropNode) { $script:DropNode.IsDropTarget = $false; $script:DropNode = $null }
    if (-not $src -or (Test-DropTarget $src $tgt)) { return }
    $script:PendingDrop = @($src, $tgt.Code)
    [void]$script:Win.Dispatcher.BeginInvoke([Action]{ $d = $script:PendingDrop; $script:PendingDrop = $null; if ($d) { Open-MoveDialog $d[0] $d[1] } }, [System.Windows.Threading.DispatcherPriority]::Background)
})

# ============================================================================ inicio
function Show-Offscreen($W) { $W.WindowStartupLocation = 'Manual'; $W.Left = -20000; $W.Top = 0; $W.ShowActivated = $false; $W.Show(); $W.UpdateLayout() }
function Run-UiTest {
    try {
        $L = $script:Lang
        $ab = Show-About -NoShow; Show-Offscreen $ab.Window; if ($ScreenshotDir) { Save-Visual $ab.Root (Join-Path $ScreenshotDir "MoveTA2_sobre_$L.png") }; $ab.Window.Close()
        Write-Host "[ui] lang=$L conectado: $($script:U.TxtConnChip.Text) planta=$($script:PlantId) TAs=$($script:Base.Count)"
        # 1) nova TA (ainda nao salva)
        $ed = New-EditDialog 'create' $script:DstIndex[$DemoCreateParent]
        $ed.C.EdCode.Text = $DemoCreate; $ed.C.EdName.Text = $DemoCreateName; Update-Edit
        Show-Offscreen $ed.Window
        if ($ScreenshotDir) { Save-Visual $ed.C.EdRoot (Join-Path $ScreenshotDir "MoveTA2_novaTA_$L.png") }
        Write-Host "[ui] nova TA: ok=$($ed.C.EdOk.IsEnabled) status='$($ed.C.EdStatusText.Text)' erro='$($ed.C.EdErr.Text)'"
        $ed.Window.Close(); Apply-Edit $ed
        # 2) mover para dentro da TA nao salva
        $why = Test-DropTarget $DemoSrc $script:DstIndex[$DemoTarget]; if ($why) { throw "drop: $why" }
        $ctx = New-MoveDialog $DemoSrc $DemoTarget
        $ctx.C.RbNext.IsChecked = $true; Update-Dialog
        Show-Offscreen $ctx.Window
        if ($ScreenshotDir) { Save-Visual $ctx.C.DlgRoot (Join-Path $ScreenshotDir "MoveTA2_dialogo_$L.png") }
        Write-Host "[ui] mover: $DemoSrc -> $($ctx.NewTop) TAs=$(@($ctx.Mapping).Count) erro='$($ctx.C.TxtErr.Text)'"
        $ctx.Window.Close(); if (-not $ctx.Mapping -or $ctx.C.TxtErr.Text) { throw 'invalid move dialog' }
        Add-PlanItem $ctx
        # 3) renomear
        $rn = New-EditDialog 'rename' $script:DstIndex[$DemoRename]
        $rn.C.EdName.Text = $DemoRenameTo; Update-Edit
        Show-Offscreen $rn.Window
        if ($ScreenshotDir) { Save-Visual $rn.C.EdRoot (Join-Path $ScreenshotDir "MoveTA2_renomear_$L.png") }
        $rn.Window.Close(); Apply-Edit $rn
        Write-Host "[ui] plano: $($script:U.TxtPlanSummary.Text)"
        # 4) simular
        Run-Plan $false
        Write-Host "[ui] simulacao: $($script:U.InfoTitle.Text) | $($script:U.InfoText.Text)"
        Write-Host "[ui] executar habilitado: $($script:U.BtnExec.IsEnabled)"
        $script:LastRows | Where-Object { $_.Acao -in 'criada','renomeada','ERRO','MANTIDA','ATENCAO','desativada' -or ($_.Tabela -in 'ST_WPWHERE2','ST_WHWHERE2' -and $_.Linhas -ne '0') } | ForEach-Object { Write-Host ("[ui]   {0,-24} {1,-22} {2,-11} {3,4}  {4}" -f $_.Par, $_.Tabela, $_.Acao, $_.Linhas, $_.Detalhe) }
        $a = New-ExecConfirmArgs; $cf = Show-Confirm @a -NoShow
        Show-Offscreen $cf.Window
        if ($ScreenshotDir) { Save-Visual $cf.Root (Join-Path $ScreenshotDir "MoveTA2_confirmar_$L.png") }
        $cf.Window.Close()
        # 5) cascata: remover a nova TA tira o movimento que depende dela
        $before = $script:Plan.Count
        Remove-PlanItem 0
        Write-Host "[ui] cascata: plano $before -> $($script:Plan.Count) item(ns) ao remover a nova TA ($($script:U.InfoText.Text))"
        Apply-Edit $ed; $ctx2 = New-MoveDialog $DemoSrc $DemoTarget; $ctx2.C.RbNext.IsChecked = $true; Update-Dialog; $ctx2.Window.Close(); Add-PlanItem $ctx2; Apply-Edit $rn
        # 6) mudar código: TA criada (7.02 -> 7.03, filho movido segue), topo movido (7.03.01 -> 7.03.05), TA salva (renumerar)
        $rc = New-EditDialog 'rename' $script:DstIndex[$DemoCreate]; $newCreate = $DemoCreateParent + '.03'
        $rc.C.EdCode.Text = $newCreate; Update-Edit
        Write-Host "[ui] recode criada: modo=$($rc.CodeMode) ok=$($rc.C.EdOk.IsEnabled) '$($rc.C.EdStatusText.Text)' erro='$($rc.C.EdErr.Text)'"
        $rc.Window.Close(); Apply-Edit $rc
        $mv = $script:Plan | Where-Object { $_.Type -eq 'move' } | Select-Object -First 1
        Write-Host "[ui]   movimento agora: alvo=$($mv.Target) novo topo=$($mv.NewTop)"
        $rt = New-EditDialog 'rename' $script:DstIndex[$mv.NewTop]; $newTop = $newCreate + '.05'
        $rt.C.EdCode.Text = $newTop; Update-Edit
        Show-Offscreen $rt.Window
        if ($ScreenshotDir) { Save-Visual $rt.C.EdRoot (Join-Path $ScreenshotDir "MoveTA2_codigo_$L.png") }
        Write-Host "[ui] recode topo movido: modo=$($rt.CodeMode) ok=$($rt.C.EdOk.IsEnabled) '$($rt.C.EdStatusText.Text)' erro='$($rt.C.EdErr.Text)'"
        $rt.Window.Close(); Apply-Edit $rt
        $rb = New-EditDialog 'rename' $script:DstIndex['701.10']
        $rb.C.EdCode.Text = '701.12'; Update-Edit
        Write-Host "[ui] renumerar salva: modo=$($rb.CodeMode) ok=$($rb.C.EdOk.IsEnabled) '$($rb.C.EdStatusText.Text)' erro='$($rb.C.EdErr.Text)'"
        $rb.Window.Close(); Apply-Edit $rb
        $lk = New-EditDialog 'rename' $script:DstIndex['701.12.05']
        Write-Host "[ui] filho de estrutura movida: modo=$($lk.CodeMode) codigo habilitado=$($lk.C.EdCode.IsEnabled) dica='$($lk.C.EdHint.Text)'"
        $lk.Window.Close()
        Write-Host "[ui] plano: $($script:U.TxtPlanSummary.Text)"
        Run-Plan $false
        Write-Host "[ui] simulacao 2: $($script:U.InfoTitle.Text) | $($script:U.InfoText.Text)"
        $script:LastRows | Where-Object { ($_.Tabela -eq 'ST_Function' -and $_.Acao -in 'criada','reativada','renomeada') -or $_.Acao -in 'ERRO','ATENCAO' -or ($_.Tabela -in 'ST_WPWHERE2','ST_WHWHERE2' -and $_.Linhas -ne '0') } | Select-Object -First 14 | ForEach-Object { Write-Host ("[ui]   {0,-24} {1,-16} {2,-10} {3,4}  {4}" -f $_.Par, $_.Tabela, $_.Acao, $_.Linhas, $_.Detalhe) }
        Show-NodeInTree $script:U.TreeDst $script:DstIndex[$newTop]
        $script:Win.UpdateLayout()
        if ($ScreenshotDir) { Save-Visual $script:U.Root (Join-Path $ScreenshotDir "MoveTA2_tela_$L.png") }
    } catch { Write-Host "[ui] ERRO: $($_.Exception.Message)`n$($_.ScriptStackTrace)" }
    $script:Win.Close()
}
$win.Add_Loaded({
    Apply-Language $script:Lang
    Switch-Panel 'move'
    Refresh-Plan
    if (-not $NoAutoConnect) {
        Set-Busy $true
        $ok = $false; try { $ok = Connect-Db } finally { Set-Busy $false }
        if (-not $ok) { Show-Info 'error' (T 'not_connected_t') $script:U.TxtConnStatus.Text }
    } else { Switch-Panel 'conn' }
    if ($UiTest) { [void]$script:Win.Dispatcher.BeginInvoke([Action]{ Run-UiTest }, [System.Windows.Threading.DispatcherPriority]::ContextIdle) }
})
if ($UiTest) { $win.WindowStartupLocation = 'Manual'; $win.Left = -20000; $win.Top = 0; $win.ShowActivated = $false }
[void]$win.ShowDialog()
