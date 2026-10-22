<#
.SYNOPSIS
    Move TA v2.0 - realocação de Technical Accounts no STAR IPS 3.62.x, com arrastar e soltar.
    Move TA v2.0 - drag-and-drop relocation of Technical Accounts in STAR IPS 3.62.x.

.DESCRIPTION
    Visual no padrão do FM Terminals (Fluent / novo Outlook). Português e inglês: botão PT/EN na
    barra de título troca o idioma na hora e a escolha fica salva.

    Fluxo: arraste uma TA (com toda a estrutura abaixo dela) da árvore "Estrutura atual" e
    solte sobre o novo grupo pai na "Nova estrutura". A tool pergunta como aplicar os códigos
    (trocar o início, próximo código livre ou escolher), mostra a prévia e adiciona ao plano.
    Simular roda numa transação com ROLLBACK; Executar grava (COMMIT).

    Usa dbo.usp_IPS_MoveTA / dbo.usp_IPS_RetireTA (usp_IPS_MoveTA.sql na mesma pasta; instala
    se faltarem). Versão anterior (WinForms): versions\v1.0.

.PARAMETER Lang
    pt ou en. Sem parâmetro usa o último idioma escolhido (padrão pt).

.PARAMETER UiTest
    Uso interno: conecta, monta um movimento de demonstração, simula, salva prints e fecha.
#>
[CmdletBinding()]
param(
    [string]$Server   = 'localhost\SQL2025',
    [string]$Database = 'ADEF_MOVETA',
    [string]$Plant    = '',
    [ValidateSet('', 'pt', 'en')][string]$Lang = '',
    [switch]$NoAutoConnect,
    [switch]$UiTest,
    [string]$ScreenshotDir,
    [string]$DemoSrc    = '701.10',
    [string]$DemoTarget = '801',
    [ValidateSet('prefix','next','custom')][string]$DemoRule = 'prefix',
    [string]$DemoCode   = ''
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Data

$script:Version      = '2.0'
$script:ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ProcFile     = Join-Path $script:ScriptDir 'usp_IPS_MoveTA.sql'
$script:LogDir       = Join-Path $script:ScriptDir 'logs'
$script:SettingsFile = Join-Path $env:APPDATA 'MoveTA\settings.json'

# ============================================================================ idiomas / languages
$script:Strings = @{
pt = @{
    search_hint = 'Pesquisar TA por código ou nome e tecle Enter'
    rail_move = 'Realocar'; rail_hist = 'Histórico'; rail_conn = 'Conexão'
    tt_min = 'Minimizar'; tt_max = 'Maximizar'; tt_close = 'Fechar'; tt_lang = 'Switch to English'
    move_title = 'Realocar Technical Accounts'
    move_sub = 'Arraste uma TA da estrutura atual (ela leva tudo que está abaixo) e solte sobre o novo grupo pai na nova estrutura.'
    plant = 'Planta'
    btn_sim = 'Simular'; tt_sim = 'Roda o plano numa transação e desfaz no fim'
    btn_exec = 'Executar'; tt_exec = 'Grava o plano (habilita depois de uma simulação sem erro)'
    btn_undo = 'Desfazer'; tt_undo = 'Desfaz o último movimento do plano (Ctrl+Z)'
    btn_clear = 'Limpar plano'; btn_reload = 'Recarregar'; tt_reload = 'Recarrega a planta do banco (F5)'; btn_opts = 'Opções'
    opt_h1 = 'O que vai junto com a TA'
    opt_always = 'Sempre: jobs, spares, counters (PFO) e where-list dos jobs (ST_WPWHERE2)'
    opt_wh = 'Histórico (ST_WorkHistory e ST_WHWHERE2)'; opt_po = 'Linhas de MR/PO (ST_POAITEM2)'
    opt_mat = 'Materiais: TA default e transações'; opt_other = 'Orçamento, custos e restrições de grupo'
    opt_h2 = 'Depois de mover'; opt_retire = 'Desativar as TAs de origem que ficarem vazias'
    opt_retclass = 'Desativar também na classe, se nenhuma outra planta usar'; opt_items = 'Listar cada objeto movido na atividade'
    src_title = 'Estrutura atual'; src_sub = 'Arraste daqui'; src_count = '{0} TAs'
    dst_title = 'Nova estrutura'; dst_hint = 'Solte sobre o grupo pai'
    chip_new = 'NOVA'; chip_react = 'REATIVAR'; chip_merge = 'MERGE'
    plan_title = 'Plano de realocação'; plan_sub = 'Executado na ordem, tudo ou nada'; tt_remove = 'Remover do plano'
    plan_empty_t = 'Nenhum movimento ainda'
    plan_empty_s = 'Arraste uma TA da estrutura atual e solte sobre o grupo de destino. Você escolhe como os códigos mudam antes de entrar no plano.'
    act_title = 'Atividade'; act_none = 'Nenhuma simulação ainda'; btn_export = 'Exportar log'
    col_move = 'Movimento'; col_table = 'Tabela'; col_action = 'Ação'; col_rows = 'Linhas'; col_detail = 'Detalhe'
    hist_title = 'Histórico'; hist_sub = 'Cada simulação e execução fica salva em CSV na pasta logs, ao lado da tool.'
    btn_open = 'Abrir'; btn_openfolder = 'Abrir pasta'; btn_refresh = 'Atualizar'
    col_date = 'Data'; col_kind = 'Tipo'; col_plant = 'Planta'; col_file = 'Arquivo'
    hist_empty = 'Nenhum log ainda. Rode uma simulação na tela Realocar.'
    conn_title = 'Conexão'; conn_sub = 'Banco do STAR IPS 3.62 onde as TAs serão realocadas.'
    lbl_server = 'Servidor'; lbl_db = 'Banco'; lbl_auth = 'Autenticação'; auth_win = 'Windows'; auth_sql = 'SQL login'
    lbl_user = 'Usuário'; lbl_pass = 'Senha'; btn_connect = 'Conectar'; btn_install = 'Reinstalar procedures'
    dlg_title = 'Mover estrutura'; dlg_from = 'DE'; dlg_to = 'PARA DENTRO DE'; dlg_how = 'Como aplicar os novos códigos?'
    rule_prefix = 'Trocar só o início do código'; rule_next = 'Usar o próximo código livre'; rule_custom = 'Escolher o código'
    btn_cancel = 'Cancelar'; btn_add = 'Adicionar ao plano'
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
    info_added_t = 'Adicionado ao plano'; info_added = '{0} → {1} com {2} TAs. Arraste mais estruturas ou clique em Simular.'
    st_out = 'FORA DO PADRÃO'; st_same = 'SEM MUDANÇA'; st_conflict = 'CONFLITO'; st_merge_plan = 'MERGE (plano)'
    st_merge = 'MERGE'; st_activate = 'ATIVAR'; st_react = 'REATIVAR'; st_new = 'NOVA'
    drag_root = 'A raiz do plano de contas não pode ser movida.'
    drag_inplan = '{0} já está no plano. Remova o movimento para mudar o destino.'
    drag_child = '{0}, que fica dentro de {1}, já está no plano. Remova-o primeiro ou mova só o resto.'
    drag_cant_t = 'Não dá para arrastar essa TA'
    drop_none = 'Solte sobre o grupo pai de destino.'
    drop_ghost = 'Solte sobre um grupo que já existe. TAs novas do plano ainda não aceitam filhos.'
    drop_leaving = '{0} está saindo no plano.'; drop_self = 'Uma TA não pode ir para dentro dela mesma.'
    drag_from = 'Arraste uma TA da estrutura atual.'
    dragging = 'Arrastando {0} com {1} TA(s): solte sobre o novo grupo pai'; drop_here = 'Soltar {0} dentro de {1} · {2}'
    plan_count = '{0} movimento(s)'; plan_sum = '{0} movimento(s) · {1} TAs'; sum_merge = ' · {0} merge'; sum_bad = ' · {0} conflito(s)'
    plan_emptysum = 'Plano vazio'; warn_conf = '{0} conflito(s)'; warn_merge = '{0} merge'
    search_t = 'Pesquisa'; search_found = "{0} TA(s) encontradas para '{1}'. Esc limpa."; search_none = "Nada encontrado para '{0}'."
    confirm_exec = "Gravar a realocação de {0} TA(s) em {1} movimento(s) na planta {2}?`n`nBanco: {3}`nTudo roda numa transação: grava tudo ou nada.`nFaça backup antes se for produção."
    confirm_exec_t = 'Move TA - executar'
    act_line = '{0} · {1} · {2} linhas de log'; act_sim = 'Simulação'; act_exec = 'Execução'
    resumo = '{0} TAs · {1} ligações PFO · {2} em ST_WPWHERE2 · {3} em ST_WHWHERE2 · {4} históricos · {5} alertas'
    fail_exec = 'Execução falhou, nada foi gravado'; fail_sim = 'Simulação falhou'
    sim_ok_t = 'Simulação OK, nada foi gravado'; sim_ok = '{0}. Revise a atividade e clique em Executar.'
    exec_ok_t = 'Realocação gravada'; exec_ok = '{0}. Log salvo em {1}.'
    m_tas = 'TAs no plano'; m_created = 'TAs criadas / reativadas'; m_pfo = 'Jobs, spares, counters (PFO)'
    m_wp = 'Where-list jobs (ST_WPWHERE2)'; m_wh = 'Where-list histórico (ST_WHWHERE2)'; m_hist = 'Histórico (ST_WorkHistory)'
    m_retired = 'TAs de origem desativadas'; m_alerts = 'Alertas'
    log_commit = 'Alterações gravadas.'; log_rollback = 'Simulação: nada foi gravado.'; log_err = '{0} (transação desfeita, nada gravado)'; log_retire = '{0} (desativar)'
    proc_missing = "As procedures usp_IPS_MoveTA / usp_IPS_RetireTA não existem em {0}.`nInstalar agora a partir de usp_IPS_MoveTA.sql?"
    proc_notinst = 'Procedures não instaladas.'; conn_ok = 'Conectado em {0} / {1} (IPS {2}). Procedures prontas.'
    conn_fail = 'Falha: {0}'; disconnected = 'Desconectado'; not_connected_t = 'Não conectou'
    plant_loaded_t = 'Planta {0} carregada'; plant_loaded = '{0} TAs na classe {1}. Arraste uma TA da esquerda e solte sobre o novo grupo pai à direita.'
    plant_fail = 'Falha ao carregar a planta'; connect_first = 'Conecte primeiro.'; reinstalled = 'Procedures reinstaladas em {0}.'
    switch_plant = 'Trocar de planta descarta o plano atual. Continuar?'
    close_confirm = 'Há movimentos no plano que não foram executados. Fechar mesmo assim?'
    log_exported = 'Log exportado'; kind_exec = 'EXECUÇÃO'; kind_sim = 'SIMULAÇÃO'
    lang_t = 'Idioma'; lang_set = 'Interface em português.'
    cf_exec_title = 'Gravar a realocação?'; cf_exec_sub = 'Tudo roda numa única transação: ou grava tudo, ou nada.'
    cf_tas = 'TAs'; cf_moves = 'Movimentos'; cf_plant = 'Planta'; cf_db = 'Banco'
    cf_backup = 'Se este banco for produção, faça backup antes. A última simulação deste plano passou sem erro.'
    cf_go = 'Executar agora'; cf_more = '+ {0} movimento(s)'
    cf_close_t = 'Fechar o Move TA?'; cf_close_btn = 'Fechar mesmo assim'
    cf_plant_t = 'Trocar de planta?'; cf_discard = 'Descartar e trocar'
    cf_proc_t = 'Instalar as procedures?'; cf_install = 'Instalar'
}
en = @{
    search_hint = 'Search a TA by code or name and press Enter'
    rail_move = 'Relocate'; rail_hist = 'History'; rail_conn = 'Connection'
    tt_min = 'Minimize'; tt_max = 'Maximize'; tt_close = 'Close'; tt_lang = 'Mudar para português'
    move_title = 'Relocate Technical Accounts'
    move_sub = 'Drag a TA from the current structure (everything below it comes along) and drop it on the new parent group in the new structure.'
    plant = 'Plant'
    btn_sim = 'Simulate'; tt_sim = 'Runs the plan in a transaction and rolls it back at the end'
    btn_exec = 'Execute'; tt_exec = 'Commits the plan (enabled after an error-free simulation)'
    btn_undo = 'Undo'; tt_undo = 'Removes the last move from the plan (Ctrl+Z)'
    btn_clear = 'Clear plan'; btn_reload = 'Reload'; tt_reload = 'Reloads the plant from the database (F5)'; btn_opts = 'Options'
    opt_h1 = 'What moves with the TA'
    opt_always = 'Always: jobs, spares, counters (PFO) and job where-lists (ST_WPWHERE2)'
    opt_wh = 'Work history (ST_WorkHistory and ST_WHWHERE2)'; opt_po = 'MR/PO lines (ST_POAITEM2)'
    opt_mat = 'Materials: default TA and transactions'; opt_other = 'Budget, costs and user group restrictions'
    opt_h2 = 'After moving'; opt_retire = 'Deactivate source TAs that end up empty'
    opt_retclass = 'Also deactivate in the function class if no other plant uses it'; opt_items = 'List every moved object in the activity'
    src_title = 'Current structure'; src_sub = 'Drag from here'; src_count = '{0} TAs'
    dst_title = 'New structure'; dst_hint = 'Drop on the parent group'
    chip_new = 'NEW'; chip_react = 'REACTIVATE'; chip_merge = 'MERGE'
    plan_title = 'Relocation plan'; plan_sub = 'Runs in order, all or nothing'; tt_remove = 'Remove from plan'
    plan_empty_t = 'No moves yet'
    plan_empty_s = 'Drag a TA from the current structure and drop it on the target group. You choose how the codes change before it joins the plan.'
    act_title = 'Activity'; act_none = 'No simulation yet'; btn_export = 'Export log'
    col_move = 'Move'; col_table = 'Table'; col_action = 'Action'; col_rows = 'Rows'; col_detail = 'Detail (from the database)'
    hist_title = 'History'; hist_sub = 'Every simulation and execution is saved as CSV in the logs folder next to the tool.'
    btn_open = 'Open'; btn_openfolder = 'Open folder'; btn_refresh = 'Refresh'
    col_date = 'Date'; col_kind = 'Type'; col_plant = 'Plant'; col_file = 'File'
    hist_empty = 'No logs yet. Run a simulation on the Relocate screen.'
    conn_title = 'Connection'; conn_sub = 'STAR IPS 3.62 database where the TAs will be relocated.'
    lbl_server = 'Server'; lbl_db = 'Database'; lbl_auth = 'Authentication'; auth_win = 'Windows'; auth_sql = 'SQL login'
    lbl_user = 'User'; lbl_pass = 'Password'; btn_connect = 'Connect'; btn_install = 'Reinstall procedures'
    dlg_title = 'Move structure'; dlg_from = 'FROM'; dlg_to = 'INTO'; dlg_how = 'How should the new codes be applied?'
    rule_prefix = 'Only change the start of the code'; rule_next = 'Use the next free code'; rule_custom = 'Choose the code'
    btn_cancel = 'Cancel'; btn_add = 'Add to plan'
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
    info_added_t = 'Added to plan'; info_added = '{0} → {1} with {2} TAs. Drag more structures or click Simulate.'
    st_out = 'OFF PATTERN'; st_same = 'NO CHANGE'; st_conflict = 'CONFLICT'; st_merge_plan = 'MERGE (plan)'
    st_merge = 'MERGE'; st_activate = 'ACTIVATE'; st_react = 'REACTIVATE'; st_new = 'NEW'
    drag_root = 'The root of the chart of accounts cannot be moved.'
    drag_inplan = '{0} is already in the plan. Remove that move to change its target.'
    drag_child = '{0}, which is inside {1}, is already in the plan. Remove it first or move only the rest.'
    drag_cant_t = 'This TA cannot be dragged'
    drop_none = 'Drop on the target parent group.'
    drop_ghost = 'Drop on a group that already exists. New TAs in the plan do not accept children yet.'
    drop_leaving = '{0} is leaving in the plan.'; drop_self = 'A TA cannot go inside itself.'
    drag_from = 'Drag a TA from the current structure.'
    dragging = 'Dragging {0} with {1} TA(s): drop on the new parent group'; drop_here = 'Drop {0} into {1} · {2}'
    plan_count = '{0} move(s)'; plan_sum = '{0} move(s) · {1} TAs'; sum_merge = ' · {0} merge'; sum_bad = ' · {0} conflict(s)'
    plan_emptysum = 'Empty plan'; warn_conf = '{0} conflict(s)'; warn_merge = '{0} merge'
    search_t = 'Search'; search_found = "{0} TA(s) found for '{1}'. Esc clears."; search_none = "Nothing found for '{0}'."
    confirm_exec = "Commit the relocation of {0} TA(s) in {1} move(s) at plant {2}?`n`nDatabase: {3}`nEverything runs in one transaction: all or nothing.`nTake a backup first if this is production."
    confirm_exec_t = 'Move TA - execute'
    act_line = '{0} · {1} · {2} log lines'; act_sim = 'Simulation'; act_exec = 'Execution'
    resumo = '{0} TAs · {1} PFO links · {2} in ST_WPWHERE2 · {3} in ST_WHWHERE2 · {4} work histories · {5} alerts'
    fail_exec = 'Execution failed, nothing was saved'; fail_sim = 'Simulation failed'
    sim_ok_t = 'Simulation OK, nothing was saved'; sim_ok = '{0}. Review the activity and click Execute.'
    exec_ok_t = 'Relocation saved'; exec_ok = '{0}. Log saved to {1}.'
    m_tas = 'TAs in plan'; m_created = 'TAs created / reactivated'; m_pfo = 'Jobs, spares, counters (PFO)'
    m_wp = 'Job where-lists (ST_WPWHERE2)'; m_wh = 'History where-lists (ST_WHWHERE2)'; m_hist = 'Work history (ST_WorkHistory)'
    m_retired = 'Source TAs deactivated'; m_alerts = 'Alerts'
    log_commit = 'Changes saved.'; log_rollback = 'Simulation: nothing was saved.'; log_err = '{0} (transaction rolled back, nothing saved)'; log_retire = '{0} (deactivate)'
    proc_missing = "The procedures usp_IPS_MoveTA / usp_IPS_RetireTA do not exist in {0}.`nInstall them now from usp_IPS_MoveTA.sql?"
    proc_notinst = 'Procedures not installed.'; conn_ok = 'Connected to {0} / {1} (IPS {2}). Procedures ready.'
    conn_fail = 'Failed: {0}'; disconnected = 'Disconnected'; not_connected_t = 'Not connected'
    plant_loaded_t = 'Plant {0} loaded'; plant_loaded = '{0} TAs in class {1}. Drag a TA from the left and drop it on the new parent group on the right.'
    plant_fail = 'Failed to load the plant'; connect_first = 'Connect first.'; reinstalled = 'Procedures reinstalled in {0}.'
    switch_plant = 'Switching plants discards the current plan. Continue?'
    close_confirm = 'There are moves in the plan that were not executed. Close anyway?'
    log_exported = 'Log exported'; kind_exec = 'EXECUTION'; kind_sim = 'SIMULATION'
    lang_t = 'Language'; lang_set = 'Interface in English. Log details come from the database procedures and stay in Portuguese.'
    cf_exec_title = 'Commit the relocation?'; cf_exec_sub = 'Everything runs in a single transaction: all or nothing.'
    cf_tas = 'TAs'; cf_moves = 'Moves'; cf_plant = 'Plant'; cf_db = 'Database'
    cf_backup = 'If this database is production, take a backup first. The last simulation of this plan passed without errors.'
    cf_go = 'Execute now'; cf_more = '+ {0} more move(s)'
    cf_close_t = 'Close Move TA?'; cf_close_btn = 'Close anyway'
    cf_plant_t = 'Switch plant?'; cf_discard = 'Discard and switch'
    cf_proc_t = 'Install the procedures?'; cf_install = 'Install'
}
}
$script:LinkLabelsByLang = @{
    pt = @{ 'jobs'=@('job','jobs'); 'WOs'=@('WO','WOs'); 'WH'=@('hist.','hist.'); 'where-list'=@('where','where'); 'spares'=@('spare','spares'); 'counters'=@('cont.','cont.'); 'equip'=@('equip.','equip.'); 'assemblies'=@('assy','assy'); 'outros'=@('outro','outros') }
    en = @{ 'jobs'=@('job','jobs'); 'WOs'=@('WO','WOs'); 'WH'=@('hist.','hist.'); 'where-list'=@('where','where'); 'spares'=@('spare','spares'); 'counters'=@('counter','counters'); 'equip'=@('equip.','equip.'); 'assemblies'=@('assy','assy'); 'outros'=@('other','others') }
}
$script:ActionEn = @{ 'inicio'='start'; 'fim'='end'; 'existe'='exists'; 'criada'='created'; 'reativada'='reactivated'; 'copiada'='copied'; 'movidas'='moved'; 'atualizadas'='updated'
                      'desativada'='deactivated'; 'MANTIDA'='KEPT'; 'ATENCAO'='WARNING'; 'VERIFICAR'='CHECK'; 'pulado'='skipped'; 'item'='item'; 'ja inativa'='already inactive'
                      'COMMIT'='COMMIT'; 'ROLLBACK'='ROLLBACK'; 'ERRO'='ERROR' }

function Get-SavedLang { try { if (Test-Path $script:SettingsFile) { return [string](Get-Content $script:SettingsFile -Raw | ConvertFrom-Json).lang } } catch {}; '' }
function Save-Lang { if ($UiTest) { return }; try { $d = Split-Path $script:SettingsFile; if (-not (Test-Path $d)) { New-Item -ItemType Directory $d -Force | Out-Null }; @{ lang = $script:Lang } | ConvertTo-Json | Set-Content $script:SettingsFile -Encoding UTF8 } catch {} }
$script:Lang = $(if ($Lang) { $Lang } else { $s = Get-SavedLang; if ($s -in 'pt','en') { $s } else { 'pt' } })
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
    public string Name { get; set; }
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
    public string RowBg { get { if (_drop) return "#CFE4FA"; if (_match) return "#FFF4CE"; if (IsGhost) return "#EBF3FC"; if (_state == "merge") return "#FFF9F5"; return "Transparent"; } }
    public string RowBorder { get { return _drop ? "#0F6CBD" : "Transparent"; } }
    public string CodeFg { get { if (_state == "moved") return "#9E9E9E"; if (IsGhost) return "#0F548C"; return "#242424"; } }
    public string NameFg { get { if (_state == "moved") return "#9E9E9E"; if (IsGhost) return "#0F548C"; return "#424242"; } }
    public string Decor { get { return _state == "moved" ? "Strikethrough" : "None"; } }
    public string LinksVis { get { return string.IsNullOrEmpty(_links) ? "Collapsed" : "Visible"; } }
    public string ChipVis { get { return string.IsNullOrEmpty(_chip) ? "Collapsed" : "Visible"; } }
    public string ChipBg { get { switch (_chipKind) { case "new": return "#0F6CBD"; case "activate": case "reactivate": return "#DFF6DD"; case "merge": return "#FFE8D6"; case "invalid": return "#FDE7E9"; default: return "#F0F0F0"; } } }
    public string ChipFg { get { switch (_chipKind) { case "new": return "#FFFFFF"; case "activate": case "reactivate": return "#0E700E"; case "merge": return "#8A3707"; case "invalid": return "#B10E1C"; default: return "#616161"; } } }
    public string Tip { get { return Code + "  " + Name + (string.IsNullOrEmpty(_links) ? "" : "\n" + _links); } }
  }
  public class PreviewRow { public string From { get; set; } public string To { get; set; } public string Name { get; set; } public string Links { get; set; } public string Status { get; set; } public string StatusBg { get; set; } public string StatusFg { get; set; } }
  public class PlanCard { public int Index { get; set; } public string From { get; set; } public string To { get; set; } public string Name { get; set; } public string Stats { get; set; } public string Rule { get; set; } public string Warn { get; set; } public string WarnVis { get; set; } public string WarnBg { get; set; } public string WarnFg { get; set; } }
  public class LogRow { public string Par { get; set; } public string Tabela { get; set; } public string Acao { get; set; } public string AcaoText { get; set; } public string Linhas { get; set; } public string Detalhe { get; set; } public string RowBg { get; set; } }
  public class Metric { public string Key { get; set; } public string Label { get; set; } public string Value { get; set; } public string Fg { get; set; } }
  public class HistRow { public string Date { get; set; } public string Kind { get; set; } public string Plant { get; set; } public string File { get; set; } public string Path { get; set; } public string KindBg { get; set; } public string KindFg { get; set; } }
}
'@
}

# ============================================================================ estado
$script:Conn = $null; $script:ConnLabel = $null; $script:ConnStatusKey = $null; $script:ConnStatusArgs = @()
$script:PlantId = $null; $script:FuncClass = $null
$script:Base = @{}; $script:ClassSet = @{}; $script:DeletedSet = @{}; $script:Kids = @{}; $script:SortedCodes = @()
$script:Plan = New-Object System.Collections.ArrayList
$script:SrcIndex = @{}; $script:DstIndex = @{}
$script:SimOk = $false; $script:LastRows = $null; $script:LastPairs = 0; $script:LastLogFile = $null; $script:LastKind = $null; $script:LastTime = $null
$script:DlgCtx = $null; $script:DragNode = $null; $script:DragStart = $null
$script:DropNode = $null; $script:HoverNode = $null; $script:HoverSince = [datetime]::MinValue
$script:PendingDrop = $null; $script:Loading = $false; $script:DstScroll = $null

function B([string]$Hex) { (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex) }

# ============================================================================ SQL
function New-ConnString { param([string]$Srv, [string]$Db, [bool]$WinAuth, [string]$User, [string]$Pass)
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source'] = $Srv; $b['Initial Catalog'] = $Db; $b['TrustServerCertificate'] = $true
    $b['Application Name'] = 'MoveTA 2.0'; $b['Connect Timeout'] = 8
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
function Test-Procs { (Invoke-SqlScalar "SELECT CASE WHEN OBJECT_ID('dbo.usp_IPS_MoveTA') IS NOT NULL AND OBJECT_ID('dbo.usp_IPS_RetireTA') IS NOT NULL THEN 1 ELSE 0 END") -eq 1 }
function Install-Procs {
    if (-not (Test-Path $script:ProcFile)) { throw "usp_IPS_MoveTA.sql not found: $($script:ProcFile)" }
    foreach ($b in [regex]::Split((Get-Content $script:ProcFile -Raw), '(?im)^\s*GO\s*$')) {
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

# ============================================================================ plano
function Get-PlannedTargets([int]$Upto = -1) {
    $h = @{}
    for ($i = 0; $i -lt $script:Plan.Count; $i++) { if ($Upto -ge 0 -and $i -ge $Upto) { break }; foreach ($p in $script:Plan[$i].Pairs) { if ($p.To) { $h[$p.To] = $p.From } } }
    $h
}
function Get-PlannedSources([int]$Skip = -1) {
    $h = @{}
    for ($i = 0; $i -lt $script:Plan.Count; $i++) { if ($i -eq $Skip) { continue }; foreach ($p in $script:Plan[$i].Pairs) { $h[$p.From] = $p.To } }
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
        [void]$rows.Add([pscustomobject]@{ From = $c; To = $to; Name = $script:Base[$c].Name; Status = $st[0]; Kind = $st[1]; Links = $script:Base[$c].LinkText; LinkH = $script:Base[$c].Links })
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
function Test-DragSource([string]$Code) {
    if ($Code -eq '*' -or -not $script:Base[$Code].Father) { return (T 'drag_root') }
    $src = Get-PlannedSources
    if ($src.ContainsKey($Code)) { return ((T 'drag_inplan') -f $Code) }
    foreach ($it in $script:Plan) { if (Test-IsAncestor $Code $it.Src) { return ((T 'drag_child') -f $it.Src, $Code) } }
    $null
}
function Test-DropTarget([string]$Src, $TgtNode) {
    if (-not $TgtNode) { return (T 'drop_none') }
    if ($TgtNode.IsGhost) { return (T 'drop_ghost') }
    if ($TgtNode.State -eq 'moved') { return ((T 'drop_leaving') -f $TgtNode.Code) }
    if (Test-IsAncestor $Src $TgtNode.Code) { return (T 'drop_self') }
    $null
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
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
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
        Title="Move TA 2.0" Width="1440" Height="900" MinWidth="1180" MinHeight="700"
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
            <TextBlock x:Name="TxtVersion" Text="v2.0" Foreground="White" FontSize="11"/>
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
        </StackPanel>

        <Grid Grid.Column="1">
          <Grid x:Name="PanelMove" Margin="4,8,12,12">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="220"/><RowDefinition Height="8"/><RowDefinition Height="240" MinHeight="120"/></Grid.RowDefinitions>
            <Grid Margin="8,4,0,10">
              <StackPanel>
                <TextBlock Text="{DynamicResource move_title}" FontSize="20" FontWeight="SemiBold"/>
                <TextBlock Text="{DynamicResource move_sub}" Foreground="#616161" FontSize="13" Margin="0,2,0,0"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                <TextBlock Text="{DynamicResource plant}" Foreground="#616161" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="13"/>
                <ComboBox x:Name="CbPlant" Width="280" Height="32" VerticalContentAlignment="Center" FontSize="13"/>
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
                      </StackPanel>
                      <StackPanel>
                        <TextBlock Text="{DynamicResource dst_title}" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock x:Name="TxtDstHint" Foreground="#616161" FontSize="12"/>
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
                                    <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="11" Margin="8,0" VerticalAlignment="Center" Foreground="#616161"/>
                                    <TextBlock Text="{Binding To}" FontWeight="SemiBold" Foreground="#0F548C"/>
                                  </StackPanel>
                                  <TextBlock Text="{Binding Name}" Foreground="#424242" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
                                  <TextBlock Text="{Binding Stats}" Foreground="#616161" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                  <WrapPanel Margin="0,6,0,0">
                                    <Border Background="#EBF3FC" CornerRadius="4" Padding="6,1"><TextBlock Text="{Binding Rule}" FontSize="11" Foreground="#0F548C"/></Border>
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
                      <Border Background="#FAFAFA" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="6" Padding="12,5" Margin="4" MinWidth="118">
                        <StackPanel><TextBlock Text="{Binding Value}" FontSize="18" FontWeight="SemiBold" Foreground="{Binding Fg}"/><TextBlock Text="{Binding Label}" FontSize="11" Foreground="#616161"/></StackPanel>
                      </Border>
                    </DataTemplate>
                  </ItemsControl.ItemTemplate>
                </ItemsControl>
                <DataGrid x:Name="GridLog" Margin="1,0,1,1">
                  <DataGrid.RowStyle><Style TargetType="DataGridRow"><Setter Property="Background" Value="{Binding RowBg}"/></Style></DataGrid.RowStyle>
                  <DataGrid.Columns>
                    <DataGridTextColumn Binding="{Binding Par}" Width="170" ElementStyle="{StaticResource CellCode}"/>
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
              <TextBlock x:Name="TxtTgtCode" FontSize="20" FontWeight="SemiBold" Foreground="#0F548C"/>
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
                      <TextBlock DockPanel.Dock="Right" Text="{Binding Stats}" Foreground="#616161" FontSize="12" VerticalAlignment="Center" Margin="12,0,0,0"/>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding From}" FontWeight="SemiBold" FontSize="13"/>
                        <TextBlock FontFamily="{StaticResource Icons}" Text="&#xE72A;" FontSize="11" Margin="8,0" VerticalAlignment="Center" Foreground="#616161"/>
                        <TextBlock Text="{Binding To}" FontWeight="SemiBold" FontSize="13" Foreground="#0F548C"/>
                        <TextBlock Text="{Binding Name}" Foreground="#616161" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" MaxWidth="230"/>
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

function Set-LangResources($Element) { foreach ($k in $script:Strings[$script:Lang].Keys) { $Element.Resources[$k] = [string]$script:Strings[$script:Lang][$k] } }
function Set-ColumnHeaders($Grid, [string[]]$Keys) { for ($i = 0; $i -lt $Keys.Count -and $i -lt $Grid.Columns.Count; $i++) { $Grid.Columns[$i].Header = T $Keys[$i] } }

# ============================================================================ janela
$win = [Windows.Markup.XamlReader]::Parse($script:MainXaml.Replace('__STYLES__', $script:StylesXaml))
$script:Win = $win
$script:U = @{}
foreach ($n in @('Root','TxtVersion','TbSearch','TxtSearchHint','DotConn','TxtConnChip','BtnLang','BtnMin','BtnMax','BtnClose','RailMove','RailHist','RailConn',
                 'PanelMove','PanelHist','PanelConn','CbPlant','BtnSim','BtnExec','BtnUndo','BtnClear','BtnReload','TgOptions','PopOptions',
                 'ChkWH','ChkPO','ChkMat','ChkOther','ChkRetire','ChkRetireClass','ChkItems','TxtPlanSummary','InfoBar','InfoIcon','InfoTitle','InfoText','BtnInfoClose',
                 'TreeSrc','TxtSrcCount','TreeDst','TxtDstHint','PlanList','PlanEmpty','TxtPlanCount','BtnSaveLog','TxtActivity','Metrics','GridLog',
                 'TbServer','TbDb','RbWin','RbSql','TbUser','PbPass','BtnConnect','BtnInstall','TxtConnStatus','GridLogs','BtnOpenLog','BtnOpenFolder','BtnRefreshLogs','TxtHistEmpty')) {
    $script:U[$n] = $win.FindName($n)
}
$U = $script:U
$U.TxtVersion.Text = "v$($script:Version)"
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
    <# Caixa de confirmação no visual da tool. Devolve $true se confirmado (ou a janela, com -NoShow). #>
    param([string]$Title, [string]$Sub, [string]$OkText, [ValidateSet('warn','question')][string]$Kind = 'question',
          $Stats = $null, $Moves = $null, [string]$More = '', [string]$Warn = '', [string]$OkGlyph = '', [switch]$NoShow)
    $w = [Windows.Markup.XamlReader]::Parse($script:ConfirmXaml.Replace('__STYLES__', $script:StylesXaml))
    Set-LangResources $w
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
function New-ExecConfirmArgs {
    $n = 0; foreach ($it in $script:Plan) { $n += @($it.Pairs).Count }
    $stats = New-Object 'System.Collections.Generic.List[MoveTA.Metric]'
    foreach ($s in @(@((T 'cf_tas'), [string]$n, '#0F548C'), @((T 'cf_moves'), [string]$script:Plan.Count, '#242424'), @((T 'cf_plant'), $script:PlantId, '#242424'), @((T 'cf_db'), $script:DbName, '#242424'))) {
        $x = New-Object MoveTA.Metric; $x.Label = $s[0]; $x.Value = $s[1]; $x.Fg = $s[2]; $stats.Add($x)
    }
    $moves = New-Object 'System.Collections.Generic.List[MoveTA.PlanCard]'
    for ($i = 0; $i -lt [Math]::Min(6, $script:Plan.Count); $i++) {
        $it = $script:Plan[$i]; $m = New-Object MoveTA.PlanCard; $m.From = $it.Src; $m.To = $it.NewTop; $m.Name = $it.Name; $m.Stats = Get-StatsText $it.Pairs; $moves.Add($m)
    }
    $more = $(if ($script:Plan.Count -gt 6) { (T 'cf_more') -f ($script:Plan.Count - 6) } else { '' })
    @{ Title = (T 'cf_exec_title'); Sub = ((T 'cf_exec_sub') + "  ·  $($script:ConnLabel)"); OkText = (T 'cf_go'); Kind = 'warn'; Stats = $stats; Moves = $moves; More = $more; Warn = (T 'cf_backup'); OkGlyph = [string][char]0xE73E }
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
            $toByFrom = @{}; foreach ($p in $it.Pairs) { $toByFrom[$p.From] = $p.To }
            foreach ($p in ($it.Pairs | Sort-Object @{ e = { $_.To.Length } }, @{ e = { $_.To } })) {
                $srcNode = $idx[$p.From]
                if ($srcNode -and -not $srcNode.IsGhost) { $srcNode.State = 'moved'; $srcNode.ChipKind = 'moved'; $srcNode.ChipText = '→ ' + $p.To }
                if (-not $p.To) { continue }
                if ($p.Kind -eq 'merge' -and $idx.ContainsKey($p.To)) {
                    $t = $idx[$p.To]
                    if (-not $t.IsGhost) { $t.State = 'merge' }
                    $t.ChipKind = 'merge'; $t.ChipText = 'MERGE ← ' + $p.From
                    continue
                }
                if ($p.Kind -eq 'invalid') { continue }
                if ($p.From -eq $it.Src) { $parentCode = $it.Target } else { $parentCode = $toByFrom[$script:Base[$p.From].Father] }
                $parent = $idx[$parentCode]
                $g = New-Object MoveTA.TaNode; $g.Code = $p.To; $g.Name = $p.Name; $g.Links = $p.Links; $g.Father = $parentCode
                $g.State = $p.Kind; $g.ChipKind = $p.Kind; $g.ChipText = $p.Status
                if ($parent) { $coll = $parent.Children } else { $coll = $roots }
                $pos = $coll.Count; for ($i = 0; $i -lt $coll.Count; $i++) { if ([string]::CompareOrdinal($coll[$i].Code, $g.Code) -gt 0) { $pos = $i; break } }
                $coll.Insert($pos, $g); $g.Parent = $parent; $idx[$p.To] = $g
            }
            $n = $idx[$it.Target]; while ($n) { $n.IsExpanded = $true; $n = $n.Parent }
            if ($idx.ContainsKey($it.NewTop)) { $idx[$it.NewTop].IsExpanded = $true }
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
    $open = @{}; foreach ($n in $script:DstIndex.Values) { if ($n.IsExpanded -and -not $n.IsGhost) { $open[$n.Code] = 1 } }
    $r = Build-Tree $true
    foreach ($c in $open.Keys) { if ($r.Index.ContainsKey($c)) { $r.Index[$c].IsExpanded = $true } }
    foreach ($n in $r.Roots) { $n.IsExpanded = $true }
    $script:DstIndex = $r.Index
    $script:U.TreeDst.ItemsSource = $r.Roots
}
function Refresh-Plan {
    for ($i = 0; $i -lt $script:Plan.Count; $i++) { $it = $script:Plan[$i]; $it.Pairs = Get-Mapping $it.Src $it.NewTop $i; $it.RuleText = T ('rule_t_' + $it.Rule) }
    $cards = New-Object 'System.Collections.Generic.List[MoveTA.PlanCard]'
    $total = 0; $bad = 0; $merges = 0
    for ($i = 0; $i -lt $script:Plan.Count; $i++) {
        $it = $script:Plan[$i]
        $m = @($it.Pairs | Where-Object { $_.Kind -eq 'merge' }).Count
        $x = @($it.Pairs | Where-Object { $_.Kind -eq 'invalid' }).Count
        $total += @($it.Pairs).Count; $bad += $x; $merges += $m
        $card = New-Object MoveTA.PlanCard
        $card.Index = $i; $card.From = $it.Src; $card.To = $it.NewTop; $card.Name = $it.Name; $card.Stats = Get-StatsText $it.Pairs; $card.Rule = $it.RuleText
        if ($x -gt 0) { $card.Warn = (T 'warn_conf') -f $x; $card.WarnBg = '#FDE7E9'; $card.WarnFg = '#B10E1C'; $card.WarnVis = 'Visible' }
        elseif ($m -gt 0) { $card.Warn = (T 'warn_merge') -f $m; $card.WarnBg = '#FFE8D6'; $card.WarnFg = '#8A3707'; $card.WarnVis = 'Visible' }
        else { $card.WarnVis = 'Collapsed' }
        $cards.Add($card)
    }
    $script:U.PlanList.ItemsSource = $cards
    $script:U.PlanEmpty.Visibility = $(if ($script:Plan.Count) { 'Collapsed' } else { 'Visible' })
    $script:U.TxtPlanCount.Text = $(if ($script:Plan.Count) { (T 'plan_count') -f $script:Plan.Count } else { '' })
    $sources = Get-PlannedSources
    foreach ($n in $script:SrcIndex.Values) {
        if ($sources.ContainsKey($n.Code)) { $n.State = 'moved'; $n.ChipKind = 'moved'; $n.ChipText = '→ ' + $sources[$n.Code] }
        elseif ($n.State -ne 'normal') { $n.State = 'normal'; $n.ChipKind = $null; $n.ChipText = $null }
    }
    Build-DstTree
    if ($script:Plan.Count) {
        $s = (T 'plan_sum') -f $script:Plan.Count, $total
        if ($merges) { $s += (T 'sum_merge') -f $merges }
        if ($bad) { $s += (T 'sum_bad') -f $bad }
        $script:U.TxtPlanSummary.Text = $s
    } else { $script:U.TxtPlanSummary.Text = T 'plan_emptysum' }
    $script:U.BtnSim.IsEnabled = ($script:Plan.Count -gt 0 -and $bad -eq 0 -and $script:Conn)
    $script:U.BtnExec.IsEnabled = ($script:SimOk -and $script:Plan.Count -gt 0 -and $bad -eq 0)
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

# ============================================================================ dialogo
function New-MoveDialog([string]$Src, [string]$Tgt) {
    $w = [Windows.Markup.XamlReader]::Parse($script:DialogXaml.Replace('__STYLES__', $script:StylesXaml))
    Set-LangResources $w
    $c = @{}
    foreach ($n in @('DlgRoot','DlgHeader','TxtSub','TxtSrcCode','TxtSrcName','TxtStats','TxtTgtCode','TxtTgtName','RbPrefix','RbNext','RbCustom','TxtPrefixDesc','TxtNextDesc',
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
    $c.TxtTgtCode.Text = $Tgt; $c.TxtTgtName.Text = $script:Base[$Tgt].Name
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
    [void]$script:Plan.Add([pscustomobject]@{ Src = $Ctx.Src; Target = $Ctx.Tgt; Rule = $Ctx.Rule; RuleText = (T ('rule_t_' + $Ctx.Rule)); NewTop = $Ctx.NewTop; Pairs = $Ctx.Mapping; Name = $script:Base[$Ctx.Src].Name })
    $script:SimOk = $false
    Refresh-Plan
    Show-Info 'info' (T 'info_added_t') ((T 'info_added') -f $Ctx.Src, $Ctx.NewTop, @($Ctx.Mapping).Count)
}
function Open-MoveDialog([string]$Src, [string]$Tgt) {
    $ctx = New-MoveDialog $Src $Tgt
    if ($ctx.Window.ShowDialog() -and $ctx.Mapping) { Add-PlanItem $ctx }
}

# ============================================================================ execucao
function Get-Options {
    @{ WH = [bool]$script:U.ChkWH.IsChecked; PO = [bool]$script:U.ChkPO.IsChecked; Mat = [bool]$script:U.ChkMat.IsChecked; Other = [bool]$script:U.ChkOther.IsChecked
       Retire = [bool]$script:U.ChkRetire.IsChecked; RetireClass = [bool]$script:U.ChkRetireClass.IsChecked; Items = [bool]$script:U.ChkItems.IsChecked }
}
function Get-ActionText([string]$Acao) { if ($script:Lang -eq 'en' -and $script:ActionEn.ContainsKey($Acao)) { $script:ActionEn[$Acao] } else { $Acao } }
function New-LogRow([string]$Par, [string]$Tabela, [string]$Acao, $Linhas, [string]$Detalhe) {
    $r = New-Object MoveTA.LogRow; $r.Par = $Par; $r.Tabela = $Tabela; $r.Acao = $Acao; $r.AcaoText = Get-ActionText $Acao; $r.Linhas = [string]$Linhas; $r.Detalhe = $Detalhe
    $r.RowBg = $(switch ($Acao) { { $_ -in 'ERRO','ATENCAO','VERIFICAR' } { '#FDF3F4' } { $_ -in 'criada','reativada','desativada','COMMIT' } { '#F1FAF1' } { $_ -in 'MANTIDA','ROLLBACK' } { '#FFF9F5' } default { '#FFFFFF' } })
    $r
}
function Invoke-Plan([bool]$Commit) {
    $opt = Get-Options
    $pairs = @(); foreach ($it in $script:Plan) { $pairs += @($it.Pairs) }
    $ordered = $pairs | Sort-Object @{ e = { $_.To.Length } }, @{ e = { $_.To } }
    $rows = New-Object 'System.Collections.Generic.List[MoveTA.LogRow]'
    $err = $null
    $tran = $script:Conn.BeginTransaction('MoveTA')
    try {
        foreach ($p in $ordered) {
            $res = Invoke-Sql "EXEC dbo.usp_IPS_MoveTA @Plant_Id, @From, @To, NULL, 1, @WH, @PO, @Mat, @Other, 0, 0, @Items, NULL, 0" @{
                '@Plant_Id' = $script:PlantId; '@From' = $p.From; '@To' = $p.To; '@WH' = [int]$opt.WH; '@PO' = [int]$opt.PO; '@Mat' = [int]$opt.Mat; '@Other' = [int]$opt.Other; '@Items' = [int]$opt.Items } $tran
            foreach ($r in $res[0].Rows) { $rows.Add((New-LogRow "$($p.From) → $($p.To)" $r.Tabela $r.Acao $r.Linhas $r.Detalhe)) }
        }
        if ($opt.Retire) {
            foreach ($p in ($pairs | Sort-Object @{ e = { -$_.From.Length } }, @{ e = { $_.From } })) {
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
    [pscustomobject]@{ Rows = $rows; Error = $err; Pairs = $pairs.Count }
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
        Load-Plant $script:PlantId; $script:SrcIndex = @{}; Build-SrcTree; $script:DstIndex = @{}
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
        $script:ConnLabel = "$db · $srv"; $script:DbName = $db
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

# ============================================================================ idioma
function Apply-Language([string]$L) {
    $script:Lang = $L
    Set-LangResources $script:Win
    Set-ColumnHeaders $script:U.GridLog @('col_move','col_table','col_action','col_rows','col_detail')
    Set-ColumnHeaders $script:U.GridLogs @('col_date','col_kind','col_plant','col_file')
    $script:U.BtnLang.Content = $(if ($L -eq 'pt') { 'EN' } else { 'PT' })
    $script:U.TxtDstHint.Text = T 'dst_hint'
    if (-not $script:Conn) { $script:U.TxtConnChip.Text = T 'disconnected' }
    if ($script:ConnStatusKey) { Set-ConnStatus $script:ConnStatusKey $script:ConnStatusArgs }
    if ($script:Base.Count) {
        Update-LinkTexts
        Build-SrcTree
        Refresh-Plan
    } else { $script:U.TxtPlanSummary.Text = T 'plan_emptysum' }
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
$U.BtnLang.Add_Click({
    $new = $(if ($script:Lang -eq 'pt') { 'en' } else { 'pt' })
    Apply-Language $new
    Show-Info 'info' (T 'lang_t') (T 'lang_set')
})
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
    if ($e.Key -eq 'Z' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control' -and -not ($e.OriginalSource -is [System.Windows.Controls.TextBox])) { if ($script:Plan.Count) { $script:Plan.RemoveAt($script:Plan.Count - 1); $script:SimOk = $false; Refresh-Plan }; $e.Handled = $true }
    elseif ($e.Key -eq 'F5' -and $script:Conn) { Select-Plant; $e.Handled = $true }
})

$U.RailMove.Add_Click({ Switch-Panel 'move' }); $U.RailHist.Add_Click({ Switch-Panel 'hist' }); $U.RailConn.Add_Click({ Switch-Panel 'conn' })
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
            # volta a seleção para a planta atual
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
$U.BtnUndo.Add_Click({ if ($script:Plan.Count) { $script:Plan.RemoveAt($script:Plan.Count - 1); $script:SimOk = $false; Refresh-Plan } })
$U.BtnClear.Add_Click({ $script:Plan.Clear(); $script:SimOk = $false; Refresh-Plan; $script:U.InfoBar.Visibility = 'Collapsed' })
$U.BtnReload.Add_Click({ if ($script:Conn) { Select-Plant } })
$U.BtnSaveLog.Add_Click({
    if (-not $script:LastLogFile) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog; $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = [IO.Path]::GetFileName($script:LastLogFile)
    if ($dlg.ShowDialog($script:Win)) { Copy-Item $script:LastLogFile $dlg.FileName -Force; Show-Info 'success' (T 'log_exported') $dlg.FileName }
})
$U.PlanList.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{ param($s, $e)
    $idx = $e.OriginalSource.Tag
    if ($null -ne $idx -and [int]$idx -lt $script:Plan.Count) { $script:Plan.RemoveAt([int]$idx); $script:SimOk = $false; Refresh-Plan }
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
function Run-UiTest {
    try {
        Write-Host "[ui] lang=$($script:Lang) conectado: $($script:U.TxtConnChip.Text) planta=$($script:PlantId) TAs=$($script:Base.Count)"
        if (-not $script:Base.ContainsKey($DemoSrc) -or -not $script:Base.ContainsKey($DemoTarget)) { throw "demo TA not found: $DemoSrc / $DemoTarget" }
        $why = Test-DropTarget $DemoSrc $script:DstIndex[$DemoTarget]; if ($why) { throw $why }
        $ctx = New-MoveDialog $DemoSrc $DemoTarget
        switch ($DemoRule) { 'prefix' { $ctx.C.RbPrefix.IsChecked = $true } 'next' { $ctx.C.RbNext.IsChecked = $true } 'custom' { $ctx.C.RbCustom.IsChecked = $true; $ctx.C.TbCustom.Text = $DemoCode } }
        Update-Dialog
        $ctx.Window.WindowStartupLocation = 'Manual'; $ctx.Window.Left = -20000; $ctx.Window.Top = 0; $ctx.Window.ShowActivated = $false
        $ctx.Window.Show(); $ctx.Window.UpdateLayout()
        if ($ScreenshotDir) { Save-Visual $ctx.C.DlgRoot (Join-Path $ScreenshotDir "MoveTA2_dialogo_$($script:Lang).png") }
        Write-Host "[ui] dialogo: regra=$($ctx.Rule) novo=$($ctx.NewTop) TAs=$(@($ctx.Mapping).Count) erro='$($ctx.C.TxtErr.Text)'"
        $ctx.Window.Close()
        if (-not $ctx.Mapping -or $ctx.C.TxtErr.Text) { throw 'invalid dialog' }
        Add-PlanItem $ctx
        Write-Host "[ui] plano: $($script:U.TxtPlanSummary.Text)"
        Run-Plan $false
        Write-Host "[ui] simulacao: $($script:U.InfoTitle.Text) | $($script:U.InfoText.Text)"
        Write-Host "[ui] executar habilitado: $($script:U.BtnExec.IsEnabled)"
        $a = New-ExecConfirmArgs; $cf = Show-Confirm @a -NoShow
        $cf.Window.WindowStartupLocation = 'Manual'; $cf.Window.Left = -20000; $cf.Window.Top = 0; $cf.Window.ShowActivated = $false
        $cf.Window.Show(); $cf.Window.UpdateLayout()
        if ($ScreenshotDir) { Save-Visual $cf.Root (Join-Path $ScreenshotDir "MoveTA2_confirmar_$($script:Lang).png") }
        $cf.Window.Close()
        $other = $(if ($script:Lang -eq 'pt') { 'en' } else { 'pt' })
        Apply-Language $other; Apply-Language $(if ($other -eq 'pt') { 'en' } else { 'pt' })
        Write-Host "[ui] troca de idioma ida e volta OK; plano ainda tem $($script:Plan.Count) movimento(s); resumo: $($script:U.TxtPlanSummary.Text)"
        Show-NodeInTree $script:U.TreeSrc $script:SrcIndex[$DemoSrc]
        Show-NodeInTree $script:U.TreeDst $script:DstIndex[$ctx.NewTop]
        $script:Win.UpdateLayout()
        if ($ScreenshotDir) { Save-Visual $script:U.Root (Join-Path $ScreenshotDir "MoveTA2_tela_$($script:Lang).png") }
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
