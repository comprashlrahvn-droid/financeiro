<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sistema de Entrada de NFs</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root {
  --bg: #0f1117;
  --surface: #1a1d27;
  --surface2: #22263a;
  --border: #2e3348;
  --accent: #3b82f6;
  --accent2: #6366f1;
  --green: #10b981;
  --red: #ef4444;
  --yellow: #f59e0b;
  --orange: #f97316;
  --text: #e2e8f0;
  --muted: #94a3b8;
  --radius: 10px;
  --font: 'Inter', system-ui, sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;font-size:14px}
a{color:var(--accent);text-decoration:none}

/* LOGIN */
#loginScreen{display:flex;align-items:center;justify-content:center;min-height:100vh;background:radial-gradient(ellipse at 40% 40%,#1e2a4a 0%,var(--bg) 70%)}
.login-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:48px 40px;width:380px;box-shadow:0 24px 64px rgba(0,0,0,.5)}
.login-logo{text-align:center;margin-bottom:32px}
.login-logo .icon{font-size:40px;display:block;margin-bottom:8px}
.login-logo h1{font-size:22px;font-weight:700;color:var(--text)}
.login-logo p{color:var(--muted);font-size:13px;margin-top:4px}
.form-group{margin-bottom:18px}
.form-group label{display:block;font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px}
.form-group input,.form-group select,.form-group textarea{width:100%;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:10px 14px;color:var(--text);font-size:14px;font-family:var(--font);transition:border-color .2s}
.form-group input:focus,.form-group select:focus,.form-group textarea:focus{outline:none;border-color:var(--accent)}
.form-group select option{background:var(--surface)}
.btn{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border-radius:8px;border:none;font-size:13px;font-weight:600;cursor:pointer;transition:all .2s;font-family:var(--font)}
.btn-primary{background:var(--accent);color:#fff}
.btn-primary:hover{background:#2563eb}
.btn-success{background:var(--green);color:#fff}
.btn-success:hover{background:#059669}
.btn-danger{background:var(--red);color:#fff}
.btn-danger:hover{background:#dc2626}
.btn-warning{background:var(--yellow);color:#000}
.btn-warning:hover{background:#d97706}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border)}
.btn-ghost:hover{background:var(--surface2);color:var(--text)}
.btn-sm{padding:6px 12px;font-size:12px}
.btn-full{width:100%;justify-content:center}
.login-error{background:#7f1d1d22;border:1px solid #ef444433;color:#fca5a5;border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:16px;display:none}

/* LAYOUT */
#app{display:none;flex-direction:column;min-height:100vh}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.topbar-left{display:flex;align-items:center;gap:16px}
.logo{font-size:16px;font-weight:700;display:flex;align-items:center;gap:8px}
.logo span{color:var(--accent)}
.nav{display:flex;gap:2px}
.nav-btn{padding:8px 14px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;color:var(--muted);border:none;background:none;font-family:var(--font);transition:all .2s;display:flex;align-items:center;gap:6px}
.nav-btn:hover{color:var(--text);background:var(--surface2)}
.nav-btn.active{color:var(--text);background:var(--surface2)}
.topbar-right{display:flex;align-items:center;gap:12px}
.user-badge{display:flex;align-items:center;gap:8px;background:var(--surface2);border-radius:8px;padding:6px 12px;font-size:13px}
.user-role{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;padding:2px 6px;border-radius:4px}
.role-admin{background:#3b82f620;color:#93c5fd}
.role-operador{background:#10b98120;color:#6ee7b7}
.role-aprovador{background:#f59e0b20;color:#fcd34d}

/* MAIN */
.main{flex:1;padding:24px;max-width:1400px;margin:0 auto;width:100%}
.page{display:none}
.page.active{display:block}

/* DASHBOARD */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:16px;margin-bottom:24px}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--accent)}
.kpi.red::before{background:var(--red)}
.kpi.green::before{background:var(--green)}
.kpi.yellow::before{background:var(--yellow)}
.kpi.orange::before{background:var(--orange)}
.kpi-label{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px}
.kpi-value{font-size:28px;font-weight:800;line-height:1}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:4px}
.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px}
.chart-title{font-size:13px;font-weight:700;margin-bottom:16px;color:var(--text)}
.chart-wrap{position:relative;height:220px}

/* ALERTS */
.alert-section{margin-bottom:16px}
.alert-header{padding:10px 16px;border-radius:8px 8px 0 0;font-weight:700;font-size:13px;display:flex;align-items:center;gap:8px}
.alert-header.red{background:#7f1d1d;color:#fca5a5}
.alert-header.yellow{background:#78350f;color:#fcd34d}
.alert-table{background:var(--surface);border:1px solid var(--border);border-radius:0 0 8px 8px;overflow:hidden}

/* TABLE */
.table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.table-toolbar{padding:16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;border-bottom:1px solid var(--border)}
.search-box{flex:1;min-width:200px;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:8px 12px;color:var(--text);font-size:13px;font-family:var(--font)}
.search-box:focus{outline:none;border-color:var(--accent)}
.tbl{width:100%;border-collapse:collapse}
.tbl th{padding:10px 14px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);background:var(--surface2);border-bottom:1px solid var(--border);white-space:nowrap}
.tbl td{padding:10px 14px;border-bottom:1px solid var(--border);font-size:13px;vertical-align:middle}
.tbl tr:last-child td{border-bottom:none}
.tbl tr:hover td{background:var(--surface2)}
.badge{display:inline-flex;align-items:center;padding:3px 8px;border-radius:20px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.03em}
.badge-pago{background:#10b98120;color:#6ee7b7}
.badge-vencida{background:#ef444420;color:#fca5a5}
.badge-avencer{background:#3b82f620;color:#93c5fd}
.badge-proximo{background:#f59e0b20;color:#fcd34d}
.badge-pendente{background:#6366f120;color:#c4b5fd}
.badge-validado{background:#10b98120;color:#6ee7b7}
.badge-aguardando{background:#f59e0b20;color:#fcd34d}
.badge-recusado{background:#ef444420;color:#fca5a5}
.tbl-actions{display:flex;gap:6px;flex-wrap:wrap}
.pagination{padding:12px 16px;display:flex;align-items:center;justify-content:space-between;border-top:1px solid var(--border);font-size:13px;color:var(--muted)}
.page-btns{display:flex;gap:4px}
.page-btn{width:32px;height:32px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;font-size:13px;font-family:var(--font)}
.page-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}

/* MODAL */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:1000;align-items:flex-start;justify-content:center;overflow-y:auto;padding:24px}
.modal-overlay.open{display:flex}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:16px;width:100%;max-width:860px;position:relative;margin:auto}
.modal-header{padding:20px 24px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}
.modal-title{font-size:16px;font-weight:700}
.modal-close{background:none;border:none;color:var(--muted);cursor:pointer;font-size:20px;padding:4px;line-height:1}
.modal-body{padding:24px}
.modal-footer{padding:16px 24px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:10px}
.section-title{font-size:12px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.05em;margin-bottom:16px;margin-top:8px}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.form-grid.cols3{grid-template-columns:1fr 1fr 1fr}
.form-grid.col1{grid-template-columns:1fr}
.span2{grid-column:span 2}
.span3{grid-column:span 3}
.parcela-block{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:16px;margin-bottom:12px}
.parcela-block h4{font-size:12px;font-weight:700;color:var(--muted);margin-bottom:12px}
.xml-drop{border:2px dashed var(--border);border-radius:8px;padding:32px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:16px}
.xml-drop:hover{border-color:var(--accent);background:#3b82f610}
.xml-drop.drag{border-color:var(--accent);background:#3b82f610}
.xml-icon{font-size:32px;margin-bottom:8px}
.xml-text{color:var(--muted);font-size:13px}
.xml-text strong{color:var(--accent)}
.tabs{display:flex;gap:2px;margin-bottom:20px;background:var(--bg);border-radius:8px;padding:4px}
.tab{padding:8px 16px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;color:var(--muted);border:none;background:none;font-family:var(--font);transition:all .2s}
.tab.active{background:var(--surface);color:var(--text)}
.tab-content{display:none}
.tab-content.active{display:block}
.info-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;background:var(--bg);border-radius:8px;padding:16px;margin-bottom:16px}
.info-item label{font-size:11px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:2px}
.info-item span{font-size:14px;font-weight:600;color:var(--text)}
.timeline{padding:8px 0}
.timeline-item{display:flex;gap:12px;padding:10px 0;border-bottom:1px solid var(--border)}
.timeline-item:last-child{border:none}
.timeline-dot{width:28px;height:28px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:12px;margin-top:2px}
.dot-success{background:#10b98120;color:#10b981}
.dot-warning{background:#f59e0b20;color:#f59e0b}
.dot-error{background:#ef444420;color:#ef4444}
.dot-info{background:#3b82f620;color:#3b82f6}
.timeline-content p{font-size:13px;color:var(--text)}
.timeline-content span{font-size:11px;color:var(--muted)}
.stat-row{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--border)}
.stat-row:last-child{border:none}
.notice{padding:12px 16px;border-radius:8px;font-size:13px;margin-bottom:16px;display:flex;align-items:flex-start;gap:8px}
.notice-info{background:#3b82f620;border:1px solid #3b82f640;color:#93c5fd}
.notice-warn{background:#f59e0b20;border:1px solid #f59e0b40;color:#fcd34d}
.notice-success{background:#10b98120;border:1px solid #10b98140;color:#6ee7b7}
.empty{text-align:center;padding:60px 20px;color:var(--muted)}
.empty .eicon{font-size:40px;margin-bottom:12px}
.filter-row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
select.filter{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:7px 12px;color:var(--text);font-size:13px;font-family:var(--font);cursor:pointer}
select.filter:focus{outline:none;border-color:var(--accent)}
.page-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px}
.page-header h2{font-size:20px;font-weight:800}
.approval-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-bottom:12px;display:flex;align-items:flex-start;justify-content:space-between;gap:16px}
.approval-info h3{font-size:15px;font-weight:700;margin-bottom:4px}
.approval-info p{font-size:13px;color:var(--muted);margin-bottom:8px}
.approval-meta{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--muted)}
.approval-meta span{display:flex;align-items:center;gap:4px}
.approval-actions{display:flex;gap:8px;flex-shrink:0}
.parcela-summary{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
.parcela-chip{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:11px}
.tooltip{position:relative;display:inline-block}
.tooltip:hover::after{content:attr(data-tip);position:absolute;bottom:125%;left:50%;transform:translateX(-50%);background:#1e293b;color:#e2e8f0;padding:5px 10px;border-radius:6px;font-size:11px;white-space:nowrap;z-index:999;border:1px solid var(--border)}

/* PRINT */
@media print{
  .topbar,.nav-btn,.btn-primary,.btn-ghost,.tbl-actions{display:none!important}
  body{background:#fff;color:#000}
  .table-wrap{border:1px solid #ccc}
  .tbl td,.tbl th{border:1px solid #ccc}
}

/* SCROLLBAR */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--bg)}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}

/* RESPONSIVE */
@media(max-width:768px){
  .form-grid{grid-template-columns:1fr}
  .span2,.span3{grid-column:span 1}
  .charts-grid{grid-template-columns:1fr}
  .kpi-grid{grid-template-columns:1fr 1fr}
  .main{padding:12px}
  .nav-btn span{display:none}
}
</style>
</head>
<body>

<!-- CONFIG (primeira vez) -->
<div id="configScreen" style="display:none;align-items:center;justify-content:center;min-height:100vh;background:radial-gradient(ellipse at 40% 40%,#1e2a4a 0%,var(--bg) 70%)">
  <div class="login-card" style="width:460px">
    <div class="login-logo">
      <span class="icon">🔗</span>
      <h1>Conectar à Planilha</h1>
      <p>Configure a URL do Apps Script (uma única vez)</p>
    </div>
    <div id="configError" class="login-error"></div>
    <div class="form-group">
      <label>URL do Aplicativo Web (Apps Script)</label>
      <input type="text" id="configUrl" placeholder="https://script.google.com/macros/s/.../exec">
    </div>
    <button class="btn btn-primary btn-full" onclick="testarConexao()">🔌 Conectar</button>
    <div style="margin-top:16px;background:var(--bg);border-radius:8px;padding:12px;font-size:11px;color:var(--muted)">
      Cole aqui a URL gerada ao implantar o <strong>Codigo_GAS_NF_Sistema.gs</strong> como Aplicativo da Web na sua planilha. Esta configuração fica salva neste navegador.
    </div>
  </div>
</div>

<!-- LOGIN -->
<div id="loginScreen" style="display:none">
  <div class="login-card">
    <div class="login-logo">
      <span class="icon">🏥</span>
      <h1>Sistema Financeiro NFs</h1>
      <p>Entrada e controle de notas fiscais</p>
    </div>
    <div id="loginError" class="login-error">Usuário ou senha incorretos.</div>
    <div class="form-group">
      <label>Usuário</label>
      <input type="text" id="loginUser" placeholder="Digite seu usuário" autocomplete="username">
    </div>
    <div class="form-group">
      <label>Senha</label>
      <input type="password" id="loginPass" placeholder="Digite sua senha" autocomplete="current-password">
    </div>
    <button class="btn btn-primary btn-full" onclick="doLogin()" id="btnDoLogin">Entrar</button>
    <div id="loginLoading" style="display:none;text-align:center;margin-top:12px;font-size:12px;color:var(--muted)">⏳ Conectando à planilha...</div>
    <div style="margin-top:16px;background:var(--bg);border-radius:8px;padding:12px;font-size:11px;color:var(--muted);display:flex;justify-content:space-between;align-items:center">
      <span>Conectado à planilha ✅</span>
      <button class="btn btn-ghost btn-sm" onclick="trocarPlanilha()" style="padding:2px 8px;font-size:10px">Alterar conexão</button>
    </div>
  </div>
</div>

<!-- APP -->
<div id="app">
  <div class="topbar">
    <div class="topbar-left">
      <div class="logo">📋 <span>NF</span>Sistema</div>
      <nav class="nav">
        <button class="nav-btn active" data-page="dashboard" onclick="goPage('dashboard',this)" id="navDashboard">📊 <span>Dashboard</span></button>
        <button class="nav-btn" data-page="notas" onclick="goPage('notas',this)">📄 <span>Notas Fiscais</span></button>
        <button class="nav-btn" data-page="aprovacao" onclick="goPage('aprovacao',this)" id="navAprov">✅ <span>Aprovação</span></button>
        <button class="nav-btn" data-page="fornecedores" onclick="goPage('fornecedores',this)">🏢 <span>Fornecedores</span></button>
        <button class="nav-btn" data-page="relatorios" onclick="goPage('relatorios',this)" id="navRelatorios">📑 <span>Relatórios</span></button>
      </nav>
    </div>
    <div class="topbar-right">
      <div id="syncBadge" style="font-size:11px;color:var(--muted);display:flex;align-items:center;gap:4px"></div>
      <div class="user-badge">
        <span id="userNameBadge">–</span>
        <span id="userRoleBadge" class="user-role">–</span>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="doLogout()">Sair</button>
    </div>
  </div>

  <div class="main">

    <!-- DASHBOARD -->
    <div id="page-dashboard" class="page active">
      <div class="page-header">
        <h2>Dashboard Financeiro</h2>
        <div style="display:flex;gap:8px">
          <select class="filter" id="dashYear" onchange="renderDashboard()">
            <option value="all">Todos os anos</option>
            <option value="2025">2025</option>
            <option value="2026" selected>2026</option>
          </select>
          <button class="btn btn-ghost btn-sm" onclick="renderDashboard()">🔄 Atualizar</button>
        </div>
      </div>
      <div class="kpi-grid" id="kpiGrid"></div>
      <div class="charts-grid">
        <div class="chart-card">
          <div class="chart-title">📈 Notas por Status</div>
          <div class="chart-wrap"><canvas id="chartStatus"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-title">💰 Valor por Conta Financeira</div>
          <div class="chart-wrap"><canvas id="chartContas"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-title">📅 Vencimentos por Mês</div>
          <div class="chart-wrap"><canvas id="chartMeses"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-title">🏢 Top 10 Fornecedores (valor)</div>
          <div class="chart-wrap"><canvas id="chartForn"></canvas></div>
        </div>
      </div>
      <div id="alertasSection"></div>
    </div>

    <!-- NOTAS -->
    <div id="page-notas" class="page">
      <div class="page-header">
        <h2>Notas Fiscais</h2>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn btn-ghost btn-sm" onclick="exportXLS()">📥 Exportar XLS</button>
          <button class="btn btn-ghost btn-sm" onclick="exportPDF()">📄 Exportar PDF</button>
          <button class="btn btn-primary" onclick="openModal('modalNF')" id="btnNovaNote">➕ Nova NF</button>
        </div>
      </div>
      <div class="table-wrap">
        <div class="table-toolbar">
          <input class="search-box" placeholder="🔍 Buscar por nº NF, fornecedor, CNPJ..." id="searchNotas" oninput="filterNotas()">
          <div class="filter-row">
            <select class="filter" id="filtYear" onchange="filterNotas()">
              <option value="">Todos os anos</option>
              <option value="2025">2025</option>
              <option value="2026">2026</option>
            </select>
            <select class="filter" id="filtStatus" onchange="filterNotas()">
              <option value="">Todos os status</option>
              <option value="PAGO">Pago</option>
              <option value="VENCIDA">Vencida</option>
              <option value="A VENCER">A Vencer</option>
              <option value="PRÓXIMO VENCIMENTO">Próximo Vencimento</option>
            </select>
            <select class="filter" id="filtFornecedor" onchange="filterNotas()">
              <option value="">Todos fornecedores</option>
            </select>
            <select class="filter" id="filtValidacao" onchange="filterNotas()">
              <option value="">Todas validações</option>
              <option value="VALIDADO">Validado</option>
              <option value="AGUARDANDO VALIDAÇÃO">Aguardando</option>
              <option value="RECUSADO">Recusado</option>
            </select>
          </div>
        </div>
        <div style="overflow-x:auto">
          <table class="tbl" id="tblNotas">
            <thead>
              <tr>
                <th>Nº NF</th>
                <th>Data Fat.</th>
                <th>Fornecedor</th>
                <th>Valor Total</th>
                <th>Parcelas</th>
                <th>Conta</th>
                <th>Validação</th>
                <th>Status</th>
                <th>Ações</th>
              </tr>
            </thead>
            <tbody id="tbodyNotas"></tbody>
          </table>
        </div>
        <div class="pagination">
          <span id="paginfoNotas"></span>
          <div class="page-btns" id="pagBtnsNotas"></div>
        </div>
      </div>
    </div>

    <!-- APROVAÇÃO -->
    <div id="page-aprovacao" class="page">
      <div class="page-header">
        <h2>Fila de Aprovação</h2>
        <div class="notice notice-info" style="margin:0">🔐 Somente aprovadores e administradores podem validar ou recusar NFs</div>
      </div>
      <div id="aprovacaoList"></div>
    </div>

    <!-- FORNECEDORES -->
    <div id="page-fornecedores" class="page">
      <div class="page-header">
        <h2>Fornecedores</h2>
        <button class="btn btn-primary" onclick="openModal('modalFornecedor')">➕ Novo Fornecedor</button>
      </div>
      <div class="table-wrap">
        <div class="table-toolbar">
          <input class="search-box" placeholder="🔍 Buscar fornecedor..." id="searchForn" oninput="filterFornecedores()">
        </div>
        <div style="overflow-x:auto">
          <table class="tbl">
            <thead><tr><th>#</th><th>Nome</th><th>CNPJ</th><th>Juros (%)</th><th>Multa (%)</th><th>Protesto (dias)</th><th>Ações</th></tr></thead>
            <tbody id="tbodyForn"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- RELATÓRIOS -->
    <div id="page-relatorios" class="page">
      <div class="page-header"><h2>Relatórios</h2></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px" id="relGrid">
        <div class="chart-card" style="cursor:pointer" onclick="gerarRelatorio('vencidas')">
          <div style="font-size:32px;margin-bottom:12px">🔴</div>
          <div style="font-size:16px;font-weight:700;margin-bottom:6px">Parcelas Vencidas</div>
          <div style="color:var(--muted);font-size:13px">Todas as parcelas com data de vencimento passada e status não pago</div>
          <div style="margin-top:16px;display:flex;gap:8px">
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioXLS('vencidas')">📥 XLS</button>
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioPDF('vencidas')">📄 PDF</button>
          </div>
        </div>
        <div class="chart-card" style="cursor:pointer" onclick="gerarRelatorio('proximos')">
          <div style="font-size:32px;margin-bottom:12px">🟡</div>
          <div style="font-size:16px;font-weight:700;margin-bottom:6px">Próximo Vencimento (15 dias)</div>
          <div style="color:var(--muted);font-size:13px">Parcelas que vencem nos próximos 15 dias</div>
          <div style="margin-top:16px;display:flex;gap:8px">
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioXLS('proximos')">📥 XLS</button>
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioPDF('proximos')">📄 PDF</button>
          </div>
        </div>
        <div class="chart-card" style="cursor:pointer" onclick="gerarRelatorio('periodo')">
          <div style="font-size:32px;margin-bottom:12px">📅</div>
          <div style="font-size:16px;font-weight:700;margin-bottom:6px">Relatório por Período</div>
          <div style="color:var(--muted);font-size:13px">Filtre notas por data de faturamento</div>
          <div style="margin-top:12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
            <div><label style="font-size:11px;color:var(--muted)">De</label><input type="date" id="relDe" class="form-group" style="width:100%;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:6px 10px;color:var(--text);font-family:var(--font)"></div>
            <div><label style="font-size:11px;color:var(--muted)">Até</label><input type="date" id="relAte" class="form-group" style="width:100%;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:6px 10px;color:var(--text);font-family:var(--font)"></div>
          </div>
          <div style="margin-top:12px;display:flex;gap:8px">
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioXLS('periodo')">📥 XLS</button>
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioPDF('periodo')">📄 PDF</button>
          </div>
        </div>
        <div class="chart-card" style="cursor:pointer" onclick="gerarRelatorio('fornecedor')">
          <div style="font-size:32px;margin-bottom:12px">🏢</div>
          <div style="font-size:16px;font-weight:700;margin-bottom:6px">Relatório por Fornecedor</div>
          <div style="color:var(--muted);font-size:13px">Resumo completo por fornecedor</div>
          <div style="margin-top:12px">
            <select class="filter" id="relFornecedor" style="width:100%" onclick="event.stopPropagation()">
              <option value="">Todos os fornecedores</option>
            </select>
          </div>
          <div style="margin-top:12px;display:flex;gap:8px">
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioXLS('fornecedor')">📥 XLS</button>
            <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();gerarRelatorioPDF('fornecedor')">📄 PDF</button>
          </div>
        </div>
      </div>
      <div id="relResultado" style="margin-top:24px"></div>
    </div>

  </div>
</div>

<!-- MODAL NF -->
<div class="modal-overlay" id="modalNF">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-title" id="modalNFTitle">Nova Nota Fiscal</div>
      <button class="modal-close" onclick="closeModal('modalNF')">✕</button>
    </div>
    <div class="modal-body">
      <div class="tabs">
        <button class="tab active" onclick="switchTab('tabManual',this)">✏️ Manual</button>
        <button class="tab" onclick="switchTab('tabXML',this)">📁 Importar XML</button>
      </div>

      <div id="tabManual" class="tab-content active">
        <div class="section-title">Dados da Nota Fiscal</div>
        <div class="form-grid">
          <div class="form-group"><label>Nº NF *</label><input type="text" id="nfNumero" placeholder="Ex: 12345"></div>
          <div class="form-group"><label>Data Faturamento *</label><input type="date" id="nfDataFat"></div>
          <div class="form-group"><label>Fornecedor *</label>
            <select id="nfFornecedor" onchange="preencherCNPJ()">
              <option value="">Selecione...</option>
            </select>
          </div>
          <div class="form-group"><label>CNPJ</label><input type="text" id="nfCNPJ" placeholder="00.000.000/0000-00"></div>
          <div class="form-group"><label>Valor Total NF (R$) *</label><input type="text" id="nfValorTotal" placeholder="0,00" oninput="calcParcelas()"></div>
          <div class="form-group"><label>Nº de Parcelas *</label>
            <select id="nfNumParcelas" onchange="gerarCamposParcelas()">
              <option value="1">1</option><option value="2">2</option><option value="3">3</option>
              <option value="4">4</option><option value="5">5</option><option value="6">6</option>
              <option value="7">7</option><option value="8">8</option>
            </select>
          </div>
          <div class="form-group"><label>Conta Financeira *</label>
            <select id="nfConta">
              <option value="">Selecione...</option>
              <option>FARMACIA</option><option>MANUTENÇÃO</option><option>INFORMATICA</option>
              <option>HIG/LIMPEZA E DESCARTAVEIS</option><option>MATERIAL DE ESCRITÓRIO</option>
              <option>ALIMENTAÇÃO</option><option>EQUIPAMENTOS</option><option>OBRAS</option>
              <option>OUTROS</option>
            </select>
          </div>
          <div class="form-group"><label>Ano Referência</label>
            <select id="nfAno">
              <option value="2025">2025</option>
              <option value="2026" selected>2026</option>
            </select>
          </div>
        </div>

        <div class="section-title" style="margin-top:16px">Parcelas</div>
        <div id="camposParcelas"></div>
      </div>

      <div id="tabXML" class="tab-content">
        <div class="notice notice-info">💡 Importe um arquivo XML de NF-e para preencher automaticamente os dados da nota fiscal.</div>
        <div class="xml-drop" id="xmlDrop" onclick="document.getElementById('xmlFile').click()" ondragover="xmlDragOver(event)" ondrop="xmlDrop2(event)" ondragleave="this.classList.remove('drag')">
          <div class="xml-icon">📁</div>
          <div class="xml-text">Clique ou arraste seu arquivo <strong>XML da NF-e</strong> aqui</div>
          <div class="xml-text" style="margin-top:4px;font-size:11px">Campos serão preenchidos automaticamente</div>
        </div>
        <input type="file" id="xmlFile" accept=".xml" style="display:none" onchange="parseXML(this)">
        <div id="xmlPreview" style="display:none">
          <div id="xmlInfoBanner" class="notice notice-success">✅ XML lido com sucesso! Revise os dados, selecione a <strong>Conta Financeira</strong> e clique em Salvar.</div>
          <div class="form-grid">
            <div class="form-group"><label>Nº NF</label><input type="text" id="xNfNumero"></div>
            <div class="form-group"><label>Data Faturamento</label><input type="date" id="xNfDataFat"></div>
            <div class="form-group"><label>Fornecedor (Emitente)</label><input type="text" id="xNfFornecedor"></div>
            <div class="form-group"><label>CNPJ Emitente</label><input type="text" id="xNfCNPJ"></div>
            <div class="form-group"><label>Valor Total NF</label><input type="text" id="xNfValor"></div>
            <div class="form-group"><label>Nº Parcelas <span id="dupInfoLabel" style="color:var(--accent);font-weight:700"></span></label>
              <select id="xNfParcelas" onchange="gerarCamposParcelasXML()">
                <option value="1">1</option><option value="2">2</option><option value="3">3</option>
                <option value="4">4</option><option value="5">5</option><option value="6">6</option>
                <option value="7">7</option><option value="8">8</option>
              </select>
            </div>
            <div class="form-group"><label>Conta Financeira <span style="color:var(--red)">*</span></label>
              <select id="xNfConta">
                <option value="">— Selecione a conta —</option>
                <option>FARMACIA</option><option>MANUTENÇÃO</option><option>INFORMATICA</option>
                <option>HIG/LIMPEZA E DESCARTAVEIS</option><option>MATERIAL DE ESCRITÓRIO</option>
                <option>ALIMENTAÇÃO</option><option>EQUIPAMENTOS</option><option>OBRAS</option>
                <option>OUTROS</option>
              </select>
            </div>
            <div class="form-group"><label>Ano Referência</label>
              <select id="xNfAno"><option value="2025">2025</option><option value="2026" selected>2026</option></select>
            </div>
          </div>
          <div id="camposParcelasXML"></div>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modalNF')">Cancelar</button>
      <button class="btn btn-primary" onclick="salvarNF()" id="btnSalvarNF">💾 Salvar NF</button>
    </div>
  </div>
</div>

<!-- MODAL DETALHE NF -->
<div class="modal-overlay" id="modalDetalheNF">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-title" id="detalheNFTitle">Detalhes da NF</div>
      <button class="modal-close" onclick="closeModal('modalDetalheNF')">✕</button>
    </div>
    <div class="modal-body" id="detalheNFBody"></div>
    <div class="modal-footer" id="detalheNFFooter"></div>
  </div>
</div>

<!-- MODAL BAIXAR PARCELA -->
<div class="modal-overlay" id="modalBaixar">
  <div class="modal" style="max-width:480px">
    <div class="modal-header">
      <div class="modal-title">Baixar Parcela</div>
      <button class="modal-close" onclick="closeModal('modalBaixar')">✕</button>
    </div>
    <div class="modal-body">
      <div id="baixarInfo" style="margin-bottom:16px;font-size:13px;color:var(--muted)"></div>
      <div class="form-group"><label>Data de Pagamento *</label><input type="date" id="baixarData"></div>
      <div class="form-group"><label>Valor Pago (R$) *</label><input type="text" id="baixarValor"></div>
      <div class="form-group"><label>Observação</label><input type="text" id="baixarObs" placeholder="Opcional"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modalBaixar')">Cancelar</button>
      <button class="btn btn-success" onclick="confirmarBaixar()">✅ Confirmar Pagamento</button>
    </div>
  </div>
</div>

<!-- MODAL FORNECEDOR -->
<div class="modal-overlay" id="modalFornecedor">
  <div class="modal" style="max-width:520px">
    <div class="modal-header">
      <div class="modal-title">Fornecedor</div>
      <button class="modal-close" onclick="closeModal('modalFornecedor')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-group"><label>Nome *</label><input type="text" id="fNome" placeholder="Nome do fornecedor"></div>
      <div class="form-group"><label>CNPJ *</label><input type="text" id="fCNPJ" placeholder="00.000.000/0000-00"></div>
      <div class="form-grid cols3">
        <div class="form-group"><label>Juros (%)</label><input type="number" id="fJuros" placeholder="0" step="0.01"></div>
        <div class="form-group"><label>Multa (%)</label><input type="number" id="fMulta" placeholder="0" step="0.01"></div>
        <div class="form-group"><label>Protesto (dias)</label><input type="number" id="fProtesto" placeholder="0"></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modalFornecedor')">Cancelar</button>
      <button class="btn btn-primary" onclick="salvarFornecedor()">💾 Salvar</button>
    </div>
  </div>
</div>

<!-- MODAL RECUSAR -->
<div class="modal-overlay" id="modalRecusar">
  <div class="modal" style="max-width:480px">
    <div class="modal-header">
      <div class="modal-title">Recusar NF</div>
      <button class="modal-close" onclick="closeModal('modalRecusar')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-group"><label>Motivo da recusa *</label>
        <textarea id="recusaMotivo" rows="4" placeholder="Descreva o motivo..."></textarea>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modalRecusar')">Cancelar</button>
      <button class="btn btn-danger" onclick="confirmarRecusa()">❌ Recusar</button>
    </div>
  </div>
</div>

<script>
// ====================================================
// DATA & STATE
// ====================================================
let currentUser = null;
let notas = [];
let fornecedores = [];
let notasPag = 1;
const PER_PAGE = 20;
let editingNFId = null;
let baixarTarget = null; // {nfId, parcIdx}
let recusaTarget = null;

const CONTAS = ['FARMACIA','MANUTENÇÃO','INFORMATICA','HIG/LIMPEZA E DESCARTAVEIS','MATERIAL DE ESCRITÓRIO','ALIMENTAÇÃO','EQUIPAMENTOS','OBRAS','OUTROS'];

// ====================================================
// CONEXÃO COM A PLANILHA (Google Apps Script)
// ====================================================
let GAS_URL = localStorage.getItem('nf_gas_url') || 'https://script.google.com/macros/s/AKfycbzYe-7i0clfNeyvpH0dXHK2ntpaGFKyrPgzelLNSXpGsQzyfHQjtVtKF8EfXH8MdJar/exec';
let authToken = localStorage.getItem('nf_token') || '';

// Chama o backend (GAS). method='GET' para ações de leitura simples,
// mas usamos POST sempre para evitar limite de tamanho de URL.
async function apiCall(acao, params) {
  params = params || {};
  params.acao = acao;
  if (authToken) params.token = authToken;

  try {
    const resp = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // evita preflight CORS
      body: JSON.stringify(params)
    });
    const data = await resp.json();
    if (data.expirado) {
      alert('Sua sessão expirou. Faça login novamente.');
      doLogout();
    }
    return data;
  } catch (e) {
    return { ok: false, erro: 'Falha de conexão com a planilha: ' + e.message };
  }
}

function setSyncBadge(status) {
  const el = document.getElementById('syncBadge');
  if (!el) return;
  if (status === 'ok')      el.innerHTML = '🟢 Conectado';
  else if (status === 'sync') el.innerHTML = '🔄 Sincronizando...';
  else                       el.innerHTML = '🔴 Offline';
}

// ====================================================
// TELA DE CONFIGURAÇÃO (URL do Apps Script)
// ====================================================
async function testarConexao() {
  const url = document.getElementById('configUrl').value.trim();
  const errEl = document.getElementById('configError');
  errEl.style.display = 'none';
  if (!url || !url.includes('script.google')) {
    errEl.textContent = 'Cole a URL completa do Apps Script (.../exec).';
    errEl.style.display = 'block';
    return;
  }
  GAS_URL = url;
  const res = await apiCall('ping', {});
  if (res.ok) {
    localStorage.setItem('nf_gas_url', url);
    document.getElementById('configScreen').style.display = 'none';
    document.getElementById('loginScreen').style.display = 'flex';
    document.getElementById('loginUser').focus();
    // Garante que as abas e usuários padrão existam
    apiCall('setup', {});
  } else {
    errEl.textContent = '❌ ' + (res.erro || 'Não foi possível conectar. Verifique a URL e as permissões de implantação.');
    errEl.style.display = 'block';
  }
}
function trocarPlanilha() {
  localStorage.removeItem('nf_gas_url');
  localStorage.removeItem('nf_token');
  GAS_URL = ''; authToken = '';
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('configScreen').style.display = 'flex';
  document.getElementById('configUrl').value = '';
}

// ====================================================
// CARREGAMENTO DE DADOS (via API)
// ====================================================
async function loadData() {
  setSyncBadge('sync');
  const [resNotas, resForn] = await Promise.all([
    apiCall('listarNotas', {}),
    apiCall('listarFornecedores', {})
  ]);
  notas = resNotas.ok ? resNotas.notas : [];
  fornecedores = resForn.ok ? resForn.fornecedores : [];
  normalizarNotas();
  setSyncBadge(resNotas.ok && resForn.ok ? 'ok' : 'erro');
}

// Recarrega só as notas (após salvar/editar/aprovar etc)
async function reloadNotas() {
  setSyncBadge('sync');
  const res = await apiCall('listarNotas', {});
  if (res.ok) notas = res.notas;
  normalizarNotas();
  setSyncBadge(res.ok ? 'ok' : 'erro');
  return res.ok;
}

// Garante um "id" único (numero_ano) e historico vazio para cada NF vinda do backend
function normalizarNotas() {
  notas.forEach(n => {
    n.id = n.numero + '_' + n.ano;
    if (!n.historico) n.historico = [];
  });
}

// ====================================================
// AUTH
// ====================================================
async function doLogin() {
  const u = document.getElementById('loginUser').value.trim();
  const p = document.getElementById('loginPass').value.trim();
  if (!u || !p) return;

  document.getElementById('btnDoLogin').disabled = true;
  document.getElementById('loginLoading').style.display = 'block';
  document.getElementById('loginError').style.display = 'none';

  const res = await apiCall('login', { usuario: u, senha: p });

  document.getElementById('btnDoLogin').disabled = false;
  document.getElementById('loginLoading').style.display = 'none';

  if (!res.ok) {
    document.getElementById('loginError').textContent = res.erro || 'Usuário ou senha incorretos.';
    document.getElementById('loginError').style.display = 'block';
    return;
  }

  authToken = res.token;
  localStorage.setItem('nf_token', authToken);
  currentUser = res.user; // {id, nome, role}

  document.getElementById('loginScreen').style.display='none';
  document.getElementById('app').style.display='flex';
  document.getElementById('userNameBadge').textContent = currentUser.nome;
  const rb = document.getElementById('userRoleBadge');
  rb.textContent = currentUser.role === 'admin' ? 'Admin' : currentUser.role === 'aprovador' ? 'Aprovador' : 'Operador';
  rb.className = 'user-role ' + (currentUser.role==='admin'?'role-admin':currentUser.role==='aprovador'?'role-aprovador':'role-operador');

  aplicarPermissoes();

  await loadData();
  populateSelects();
  renderNotas();
  renderFornecedores();

  // Dashboard só para admin — operador/aprovador caem direto em Notas
  if (currentUser.role === 'admin') {
    renderDashboard();
    goPage('dashboard', document.getElementById('navDashboard'));
  } else {
    goPage('notas', document.querySelector('[data-page="notas"]'));
  }
}
document.getElementById('loginPass').addEventListener('keydown', e => { if(e.key==='Enter') doLogin(); });
document.getElementById('loginUser').addEventListener('keydown', e => { if(e.key==='Enter') document.getElementById('loginPass').focus(); });

// Ajusta visibilidade de menus conforme o perfil do usuário
function aplicarPermissoes() {
  const role = currentUser.role;

  // Dashboard — somente admin
  const navDash = document.getElementById('navDashboard');
  navDash.style.display = (role === 'admin') ? '' : 'none';

  // Aprovação — operador não vê
  document.getElementById('navAprov').style.display = (role === 'operador') ? 'none' : '';

  // Relatórios — somente admin (envolve dados financeiros completos)
  document.getElementById('navRelatorios').style.display = (role === 'admin') ? '' : 'none';

  // Botão "Nova NF" — operador e admin podem lançar
  const btnNova = document.getElementById('btnNovaNote');
  if (btnNova) btnNova.style.display = (role === 'aprovador') ? 'none' : '';
}

function doLogout() {
  currentUser = null;
  authToken = '';
  localStorage.removeItem('nf_token');
  document.getElementById('loginScreen').style.display='flex';
  document.getElementById('app').style.display='none';
  document.getElementById('loginUser').value='';
  document.getElementById('loginPass').value='';
}

// ====================================================
// NAVIGATION
// ====================================================
function goPage(id, btn) {
  // Bloqueia acesso ao dashboard/relatórios para não-admins (mesmo via teclado/histórico)
  if ((id === 'dashboard' || id === 'relatorios') && currentUser.role !== 'admin') {
    id = 'notas';
    btn = document.querySelector('[data-page="notas"]');
  }
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  if(btn) btn.classList.add('active');
  if(id==='aprovacao') renderAprovacao();
  if(id==='relatorios') populateRelFornSelect();
  if(id==='dashboard') renderDashboard();
}

// ====================================================
// HELPERS
// ====================================================
function uid() { return Date.now().toString(36)+Math.random().toString(36).slice(2); }
function fmtDate(d) {
  if(!d) return '—';
  if(typeof d === 'string' && d.includes('-') && d.length===10) {
    const [y,m,dd]=d.split('-'); return `${dd}/${m}/${y}`;
  }
  const dt = new Date(d);
  if(isNaN(dt)) return d;
  return dt.toLocaleDateString('pt-BR');
}
function fmtMoney(v) {
  if(v===undefined||v===null||v==='') return '—';
  const n = parseFloat(String(v).replace(/[^\d,.-]/g,'').replace(',','.'));
  if(isNaN(n)) return v;
  return n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
}
function parseMoney(s) {
  if(!s) return 0;
  return parseFloat(String(s).replace(/[^\d,.-]/g,'').replace(',','.')) || 0;
}
function today() { return new Date().toISOString().slice(0,10); }
function calcStatus(nf) {
  const now = new Date(); now.setHours(0,0,0,0);
  const em15 = new Date(now); em15.setDate(em15.getDate()+15);
  let allPago=true, anyVencida=false, anyProximo=false;
  for(const p of (nf.parcelas||[])) {
    if(!p.vencimento) continue;
    const v = new Date(p.vencimento); v.setHours(0,0,0,0);
    const pago = String(p.situacao||'').toUpperCase()==='PAGO';
    if(!pago) {
      allPago=false;
      if(v < now) anyVencida=true;
      else if(v >= now && v <= em15) anyProximo=true;
    }
  }
  if(allPago) return 'PAGO';
  if(anyVencida) return 'VENCIDA';
  if(anyProximo) return 'PRÓXIMO VENCIMENTO';
  return 'A VENCER';
}
function statusBadge(s) {
  const m = {
    'PAGO':'badge-pago','VENCIDA':'badge-vencida','A VENCER':'badge-avencer',
    'PRÓXIMO VENCIMENTO':'badge-proximo','PENDENTE':'badge-pendente'
  };
  return `<span class="badge ${m[s]||'badge-avencer'}">${s||'—'}</span>`;
}
function validBadge(s) {
  s = String(s||'');
  const m = {'VALIDADO':'badge-validado','AGUARDANDO VALIDAÇÃO':'badge-aguardando','RECUSADO':'badge-recusado'};
  if(s.indexOf('AGUARDANDO')!==-1) return `<span class="badge badge-aguardando">AGUARDANDO VALIDAÇÃO</span>`;
  if(s.indexOf('RECUSADO')!==-1) return `<span class="badge badge-recusado">RECUSADO</span>`;
  if(s.indexOf('VALIDADO')!==-1) return `<span class="badge badge-validado">VALIDADO</span>`;
  return `<span class="badge badge-aguardando">${s||'AGUARDANDO'}</span>`;
}
function canEdit() { return currentUser && (currentUser.role==='admin'||currentUser.role==='operador'); }
function canApprove() { return currentUser && (currentUser.role==='admin'||currentUser.role==='aprovador'); }

// ====================================================
// SELECTS
// ====================================================
function populateSelects() {
  const selects = ['nfFornecedor','filtFornecedor','relFornecedor'];
  selects.forEach(id => {
    const el = document.getElementById(id);
    if(!el) return;
    const cur = el.value;
    const isFilter = id.startsWith('filt')||id.startsWith('rel');
    el.innerHTML = isFilter ? '<option value="">Todos os fornecedores</option>' : '<option value="">Selecione...</option>';
    fornecedores.sort((a,b)=>a.nome.localeCompare(b.nome)).forEach(f=>{
      el.innerHTML += `<option value="${f.nome}">${f.nome}</option>`;
    });
    if(cur) el.value=cur;
  });
}
function populateRelFornSelect() {
  const el = document.getElementById('relFornecedor');
  if(!el) return;
  el.innerHTML = '<option value="">Todos os fornecedores</option>';
  fornecedores.sort((a,b)=>a.nome.localeCompare(b.nome)).forEach(f=>{
    el.innerHTML += `<option value="${f.nome}">${f.nome}</option>`;
  });
}
function preencherCNPJ() {
  const nome = document.getElementById('nfFornecedor').value;
  const f = fornecedores.find(x=>x.nome===nome);
  if(f) document.getElementById('nfCNPJ').value = f.cnpj||'';
}

// ====================================================
// PARCELAS CAMPOS
// ====================================================
function gerarCamposParcelas() {
  const n = parseInt(document.getElementById('nfNumParcelas').value)||1;
  const total = parseMoney(document.getElementById('nfValorTotal').value)||0;
  const valParcela = n ? (total/n).toFixed(2).replace('.',',') : '';
  let html='';
  for(let i=1;i<=n;i++){
    html+=`<div class="parcela-block"><h4>Parcela ${i}</h4>
    <div class="form-grid cols3">
      <div class="form-group"><label>Vencimento *</label><input type="date" id="parc${i}_venc"></div>
      <div class="form-group"><label>Valor (R$) *</label><input type="text" id="parc${i}_val" value="${valParcela}" placeholder="0,00"></div>
      <div class="form-group"><label>Situação</label>
        <select id="parc${i}_sit">
          <option value="A VENCER">A Vencer</option>
          <option value="PRÓXIMO VENCIMENTO">Próximo Venc.</option>
          <option value="PAGO">Pago</option>
          <option value="VENCIDA">Vencida</option>
        </select>
      </div>
    </div></div>`;
  }
  document.getElementById('camposParcelas').innerHTML=html;
}
function gerarCamposParcelasXML() {
  const n = parseInt(document.getElementById('xNfParcelas').value)||1;
  const total = parseMoney(document.getElementById('xNfValor').value)||0;
  const valParcela = n ? (total/n).toFixed(2).replace('.',',') : '';
  let html='<div class="section-title" style="margin-top:16px">Parcelas</div>';
  for(let i=1;i<=n;i++){
    html+=`<div class="parcela-block"><h4>Parcela ${i}</h4>
    <div class="form-grid cols3">
      <div class="form-group"><label>Vencimento *</label><input type="date" id="xparc${i}_venc"></div>
      <div class="form-group"><label>Valor (R$) *</label><input type="text" id="xparc${i}_val" value="${valParcela}"></div>
      <div class="form-group"><label>Situação</label>
        <select id="xparc${i}_sit"><option>A VENCER</option><option>PRÓXIMO VENCIMENTO</option><option>PAGO</option><option>VENCIDA</option></select>
      </div>
    </div></div>`;
  }
  document.getElementById('camposParcelasXML').innerHTML=html;
}
function calcParcelas() {
  const n = parseInt(document.getElementById('nfNumParcelas').value)||1;
  const total = parseMoney(document.getElementById('nfValorTotal').value)||0;
  if(!total) return;
  for(let i=1;i<=n;i++){
    const el=document.getElementById(`parc${i}_val`);
    if(el) el.value = (total/n).toFixed(2).replace('.',',');
  }
}

// ====================================================
// XML PARSER
// ====================================================
function xmlDragOver(e){e.preventDefault();document.getElementById('xmlDrop').classList.add('drag')}
function xmlDrop2(e){
  e.preventDefault();
  document.getElementById('xmlDrop').classList.remove('drag');
  const file=e.dataTransfer.files[0];
  if(file && file.name.endsWith('.xml')) readXML(file);
}
function parseXML(input){
  const file=input.files[0];
  if(file) readXML(file);
}
function readXML(file){
  const reader=new FileReader();
  // Tenta UTF-8 primeiro; se o XML declarar ISO-8859-1, relê com esse encoding
  reader.onload=e=>{
    let content = e.target.result;
    // Detecta declaração de encoding no XML
    const encMatch = content.match(/encoding=["']([^"']+)["']/i);
    const declaredEnc = encMatch ? encMatch[1].toUpperCase() : 'UTF-8';
    if(declaredEnc.includes('ISO') || declaredEnc.includes('LATIN')){
      // Relê com ISO-8859-1
      const r2 = new FileReader();
      r2.onload = e2 => processXMLContent(e2.target.result);
      r2.onerror = ()=>alert('Erro ao ler o arquivo XML.');
      r2.readAsText(file, 'ISO-8859-1');
      return;
    }
    processXMLContent(content);
  };
  reader.onerror = ()=>alert('Erro ao ler o arquivo XML.');
  reader.readAsText(file, 'UTF-8');
}

function processXMLContent(content){
    try {
      const parser=new DOMParser();
      const doc=parser.parseFromString(content,'text/xml');

      // Verifica erro de parse
      const parseErr = doc.querySelector('parsererror');
      if(parseErr){ alert('XML inválido ou malformado. Verifique o arquivo.'); return; }

      // Helper: pega o texto da primeira tag encontrada (ignora namespace)
      const g=(tag)=>{const el=doc.getElementsByTagNameNS('*',tag)[0]; return el?el.textContent.trim():'';};
      // Helper: pega todos os elementos de uma tag
      const gAll=(tag)=>Array.from(doc.getElementsByTagNameNS('*',tag));

      // ── Dados básicos ──────────────────────────────────────
      document.getElementById('xNfNumero').value = g('nNF');

      // Data de emissão — pode ser dEmi (NF-e 2.x) ou dhEmi (NF-e 3.x/4.x)
      const demi = g('dEmi') || g('dhEmi') || '';
      if(demi) document.getElementById('xNfDataFat').value = demi.slice(0,10);

      // Emitente — tenta xNome, xFant
      document.getElementById('xNfFornecedor').value = g('xNome') || g('xFant') || '';

      // CNPJ emitente (primeiro CNPJ do emit)
      const emitEl = doc.getElementsByTagNameNS('*','emit')[0];
      const cnpjRaw = emitEl ? (emitEl.getElementsByTagNameNS('*','CNPJ')[0]||{}).textContent||'' : g('CNPJ');
      if(cnpjRaw) document.getElementById('xNfCNPJ').value = cnpjRaw.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/,'$1.$2.$3/$4-$5');

      // Valor total — vNF dentro de ICMSTot, ou vProd, ou vNF global
      const icms = doc.getElementsByTagNameNS('*','ICMSTot')[0];
      const vNF = (icms ? (icms.getElementsByTagNameNS('*','vNF')[0]||{}).textContent : '') || g('vNF') || g('vProd') || '';
      document.getElementById('xNfValor').value = vNF;

      // ── Duplicatas / Parcelas ──────────────────────────────
      // NF-e armazena parcelas dentro de <cobr><dup>
      // Cada <dup> tem: <nDup> número, <dVenc> vencimento, <vDup> valor
      const dups = gAll('dup');

      if(dups.length > 0){
        // Preenche número de parcelas e gera campos
        document.getElementById('xNfParcelas').value = String(dups.length);
        gerarCamposParcelasXML();
        // Mostra quantas duplicatas foram encontradas
        const lbl = document.getElementById('dupInfoLabel');
        if(lbl) lbl.textContent = `(${dups.length} duplicata${dups.length>1?'s':''} no XML)`;

        // Aguarda o DOM gerar os inputs antes de preencher
        setTimeout(()=>{
          dups.forEach((dup, idx)=>{
            const i = idx + 1;
            const dVenc = (dup.getElementsByTagNameNS('*','dVenc')[0]||{}).textContent||'';
            const vDup  = (dup.getElementsByTagNameNS('*','vDup')[0]||{}).textContent||'';
            const vencEl = document.getElementById(`xparc${i}_venc`);
            const valEl  = document.getElementById(`xparc${i}_val`);
            if(vencEl && dVenc) vencEl.value = dVenc.slice(0,10); // YYYY-MM-DD
            if(valEl  && vDup)  valEl.value  = vDup;
          });
        }, 80);
      } else {
        // Sem duplicatas — avisa o usuário
        const lbl = document.getElementById('dupInfoLabel');
        if(lbl) lbl.textContent = '(sem duplicatas no XML)';
        document.getElementById('xNfParcelas').value = '1';
        gerarCamposParcelasXML();
        setTimeout(()=>{
          const vencEl = document.getElementById('xparc1_venc');
          const valEl  = document.getElementById('xparc1_val');
          if(vencEl && demi) vencEl.value = demi.slice(0,10);
          if(valEl  && vNF)  valEl.value  = vNF;
        }, 80);
      }

      document.getElementById('xmlPreview').style.display='block';

    } catch(err){ alert('Erro ao processar o XML: '+err.message+'\n\nVerifique se o arquivo é um XML de NF-e válido.'); }
}

// ====================================================
// MODAL
// ====================================================
function openModal(id) {
  document.getElementById(id).classList.add('open');
  if(id==='modalNF' && !editingNFId) {
    document.getElementById('nfNumParcelas').value='1';
    gerarCamposParcelas();
    document.getElementById('xmlPreview').style.display='none';
    document.getElementById('tabManual').classList.add('active');
    document.getElementById('tabXML').classList.remove('active');
    document.querySelectorAll('#modalNF .tab').forEach((t,i)=>t.classList.toggle('active',i===0));
  }
}
function closeModal(id) {
  document.getElementById(id).classList.remove('open');
  if(id==='modalNF') { editingNFId=null; document.getElementById('modalNFTitle').textContent='Nova Nota Fiscal'; resetNFForm(); }
}
function resetNFForm(){
  ['nfNumero','nfCNPJ','nfValorTotal','nfDataFat'].forEach(i=>{const el=document.getElementById(i);if(el)el.value='';});
  ['nfFornecedor','nfConta','nfAno'].forEach(i=>{const el=document.getElementById(i);if(el&&el.tagName==='SELECT')el.selectedIndex=0;});
  document.getElementById('nfNumParcelas').value='1';
  gerarCamposParcelas();
}
function switchTab(tabId, btn) {
  document.querySelectorAll('#modalNF .tab-content').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('#modalNF .tab').forEach(t=>t.classList.remove('active'));
  document.getElementById(tabId).classList.add('active');
  btn.classList.add('active');
}

// ====================================================
// SALVAR NF
// ====================================================
async function salvarNF() {
  const isXML = document.getElementById('tabXML').classList.contains('active') && document.getElementById('xmlPreview').style.display!=='none';
  const get = id => { const el=document.getElementById(id); return el ? el.value.trim() : ''; };

  // IDs corretos para cada aba
  const numero     = isXML ? get('xNfNumero')    : get('nfNumero');
  const dataFat    = isXML ? get('xNfDataFat')   : get('nfDataFat');
  const fornecedor = isXML ? get('xNfFornecedor'): get('nfFornecedor');
  const cnpj       = isXML ? get('xNfCNPJ')      : get('nfCNPJ');
  const valorTotal = isXML ? get('xNfValor')      : get('nfValorTotal');
  const numParcelas= parseInt(isXML ? get('xNfParcelas') : get('nfNumParcelas'))||1;
  const conta      = isXML ? get('xNfConta')     : get('nfConta');
  const ano        = (isXML ? get('xNfAno')      : get('nfAno')) || '2026';

  const faltando = [];
  if(!numero)     faltando.push('Nº NF');
  if(!dataFat)    faltando.push('Data de Faturamento');
  if(!fornecedor) faltando.push('Fornecedor');
  if(!valorTotal) faltando.push('Valor Total');
  if(!conta)      faltando.push('Conta Financeira');
  if(faltando.length){
    alert('Preencha os campos obrigatórios:\n• ' + faltando.join('\n• '));
    return;
  }

  const pfx = isXML ? 'xparc' : 'parc';
  const parcelas=[];
  for(let i=1;i<=numParcelas;i++){
    const v  = get(`${pfx}${i}_venc`);
    const va = get(`${pfx}${i}_val`);
    const s  = get(`${pfx}${i}_sit`);
    if(!v || !va){ alert(`Preencha vencimento e valor da parcela ${i}.`); return; }
    parcelas.push({vencimento:v, valor:va, situacao:s||'A VENCER', dataPagamento:'', valorPago:''});
  }

  const nfData = {numero,dataFat,fornecedor,cnpj,valorTotal,numParcelas,conta,ano,parcelas};
  nfData.status = calcStatus(nfData);
  nfData.validacao = editingNFId ? (notas.find(x=>x.id===editingNFId)||{}).validacao || 'AGUARDANDO VALIDAÇÃO' : 'AGUARDANDO VALIDAÇÃO';

  const btn = document.getElementById('btnSalvarNF');
  btn.disabled = true; btn.textContent = '⏳ Salvando...';

  const res = await apiCall('salvarNota', { nota: JSON.stringify(nfData) });

  btn.disabled = false; btn.textContent = '💾 Salvar NF';

  if (!res.ok) { alert('Erro ao salvar: ' + (res.erro || 'desconhecido')); return; }

  await reloadNotas();
  closeModal('modalNF');
  renderNotas();
  if (currentUser.role === 'admin') renderDashboard();
  editingNFId=null;
}

// ====================================================
// RENDER NOTAS
// ====================================================
let filteredNotas = [];
function filterNotas() {
  const q = (document.getElementById('searchNotas').value||'').toLowerCase();
  const yr = document.getElementById('filtYear').value;
  const st = document.getElementById('filtStatus').value;
  const fn = document.getElementById('filtFornecedor').value;
  const vl = document.getElementById('filtValidacao').value;
  filteredNotas = notas.filter(n=>{
    if(yr && n.ano!==yr) return false;
    if(st && calcStatus(n)!==st) return false;
    if(fn && n.fornecedor!==fn) return false;
    if(vl && n.validacao!==vl) return false;
    if(q && !(`${n.numero} ${n.fornecedor} ${n.cnpj} ${n.conta}`).toLowerCase().includes(q)) return false;
    return true;
  });
  notasPag=1;
  renderNotasTable();
}
function renderNotas() {
  filteredNotas = [...notas];
  notasPag=1;
  renderNotasTable();
  populateSelects();
}
function renderNotasTable() {
  const total=filteredNotas.length;
  const pages=Math.ceil(total/PER_PAGE)||1;
  if(notasPag>pages) notasPag=pages;
  const slice=filteredNotas.slice((notasPag-1)*PER_PAGE, notasPag*PER_PAGE);
  const tb=document.getElementById('tbodyNotas');
  if(!total){tb.innerHTML=`<tr><td colspan="9" class="empty"><div class="eicon">📄</div><div>Nenhuma NF encontrada</div></td></tr>`;document.getElementById('paginfoNotas').textContent='';document.getElementById('pagBtnsNotas').innerHTML='';return;}
  tb.innerHTML=slice.map(n=>{
    const st=calcStatus(n);
    const canAct=canEdit();
    const canApr=canApprove();
    return `<tr>
      <td><strong>#${n.numero}</strong></td>
      <td>${fmtDate(n.dataFat)}</td>
      <td>${n.fornecedor}</td>
      <td>${fmtMoney(parseMoney(n.valorTotal))}</td>
      <td style="font-size:11px;color:var(--muted)">${n.numParcelas}x</td>
      <td style="font-size:12px">${n.conta||'—'}</td>
      <td>${validBadge(n.validacao)}</td>
      <td>${statusBadge(st)}</td>
      <td>
        <div class="tbl-actions">
          <button class="btn btn-ghost btn-sm" onclick="verDetalheNF('${n.id}')">👁 Ver</button>
          ${canAct?`<button class="btn btn-ghost btn-sm" onclick="editarNF('${n.id}')">✏️</button>`:''}
          ${canAct?`<button class="btn btn-danger btn-sm" onclick="excluirNF('${n.id}')">🗑</button>`:''}
        </div>
      </td>
    </tr>`;
  }).join('');
  document.getElementById('paginfoNotas').textContent=`${(notasPag-1)*PER_PAGE+1}–${Math.min(notasPag*PER_PAGE,total)} de ${total}`;
  const pb=document.getElementById('pagBtnsNotas');
  pb.innerHTML='';
  for(let i=1;i<=pages;i++){
    if(pages>10 && i>3 && i<pages-2 && Math.abs(i-notasPag)>1){if(i===4) pb.innerHTML+='<span style="padding:0 4px;color:var(--muted)">…</span>';continue;}
    const b=document.createElement('button');
    b.className='page-btn'+(i===notasPag?' active':'');
    b.textContent=i;
    b.onclick=(function(p){return function(){notasPag=p;renderNotasTable()}})(i);
    pb.appendChild(b);
  }
}

// ====================================================
// DETALHE NF
// ====================================================
function verDetalheNF(id) {
  const nf=notas.find(x=>x.id===id);
  if(!nf) return;
  const st=calcStatus(nf);
  let parcelasHTML=nf.parcelas.map((p,i)=>{
    const venc=new Date(p.vencimento);venc.setHours(0,0,0,0);
    const now=new Date();now.setHours(0,0,0,0);
    const vencida=venc<now&&p.situacao!=='PAGO';
    return `<tr>
      <td style="font-weight:700">${i+1}</td>
      <td>${fmtDate(p.vencimento)}</td>
      <td>${fmtMoney(parseMoney(p.valor))}</td>
      <td>${p.dataPagamento?fmtDate(p.dataPagamento):'—'}</td>
      <td>${p.valorPago?fmtMoney(parseMoney(p.valorPago)):'—'}</td>
      <td>${statusBadge(p.situacao||'A VENCER')}</td>
      <td>${canEdit()&&p.situacao!=='PAGO'?`<button class="btn btn-success btn-sm" onclick="abrirBaixar('${id}',${i})">💰 Baixar</button>`:''}</td>
    </tr>`;
  }).join('');
  document.getElementById('detalheNFTitle').textContent=`NF #${nf.numero} — ${nf.fornecedor}`;
  document.getElementById('detalheNFBody').innerHTML=`
    <div class="info-row" style="grid-template-columns:1fr 1fr 1fr 1fr">
      <div class="info-item"><label>Nº NF</label><span>#${nf.numero}</span></div>
      <div class="info-item"><label>Faturamento</label><span>${fmtDate(nf.dataFat)}</span></div>
      <div class="info-item"><label>Fornecedor</label><span>${nf.fornecedor}</span></div>
      <div class="info-item"><label>CNPJ</label><span>${nf.cnpj||'—'}</span></div>
      <div class="info-item"><label>Valor Total</label><span>${fmtMoney(parseMoney(nf.valorTotal))}</span></div>
      <div class="info-item"><label>Conta</label><span>${nf.conta||'—'}</span></div>
      <div class="info-item"><label>Validação</label><span>${validBadge(nf.validacao)}</span></div>
      <div class="info-item"><label>Status</label><span>${statusBadge(st)}</span></div>
    </div>
    <div class="section-title">Parcelas</div>
    <div style="overflow-x:auto">
    <table class="tbl">
      <thead><tr><th>#</th><th>Vencimento</th><th>Valor</th><th>Data Pag.</th><th>Valor Pago</th><th>Status</th><th>Ação</th></tr></thead>
      <tbody>${parcelasHTML}</tbody>
    </table></div>
  `;
  const footer=document.getElementById('detalheNFFooter');
  footer.innerHTML=`<button class="btn btn-ghost" onclick="closeModal('modalDetalheNF')">Fechar</button>`;
  if((nf.validacao||'').indexOf('RECUSADO')!==-1){
    const motivo = nf.validacao.split(':').slice(1).join(':').trim();
    footer.innerHTML=`<div class="notice notice-warn" style="flex:1">❌ Recusado: ${motivo||'sem motivo informado'}</div>`+footer.innerHTML;
  }
  openModal('modalDetalheNF');
}

// ====================================================
// BAIXAR PARCELA
// ====================================================
function abrirBaixar(nfId, idx) {
  baixarTarget={nfId,idx};
  const nf=notas.find(x=>x.id===nfId);
  const p=nf.parcelas[idx];
  document.getElementById('baixarInfo').innerHTML=`<strong>NF #${nf.numero}</strong> — ${nf.fornecedor}<br>Parcela ${idx+1} | Vencimento: ${fmtDate(p.vencimento)} | Valor: ${fmtMoney(parseMoney(p.valor))}`;
  document.getElementById('baixarData').value=today();
  document.getElementById('baixarValor').value=String(p.valor||'').replace('.',',');
  document.getElementById('baixarObs').value='';
  openModal('modalBaixar');
}
async function confirmarBaixar() {
  const data=document.getElementById('baixarData').value;
  const valor=document.getElementById('baixarValor').value;
  if(!data||!valor){alert('Preencha data e valor.');return;}
  const nf=notas.find(x=>x.id===baixarTarget.nfId);

  const res = await apiCall('baixarParcela', {
    numero: nf.numero, ano: nf.ano, parcIdx: baixarTarget.idx,
    dataPagamento: data, valorPago: valor
  });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }

  await reloadNotas();
  closeModal('modalBaixar');
  closeModal('modalDetalheNF');
  renderNotas();
  if (currentUser.role === 'admin') renderDashboard();
  setTimeout(()=>verDetalheNF(baixarTarget.nfId),100);
}

// ====================================================
// EDITAR / EXCLUIR
// ====================================================
function editarNF(id) {
  const nf=notas.find(x=>x.id===id);
  if(!nf) return;
  editingNFId=id;
  document.getElementById('modalNFTitle').textContent=`Editar NF #${nf.numero}`;
  document.getElementById('nfNumero').value=nf.numero;
  document.getElementById('nfDataFat').value=nf.dataFat;
  document.getElementById('nfFornecedor').value=nf.fornecedor;
  document.getElementById('nfCNPJ').value=nf.cnpj||'';
  document.getElementById('nfValorTotal').value=nf.valorTotal;
  document.getElementById('nfNumParcelas').value=nf.numParcelas||nf.parcelas.length;
  document.getElementById('nfConta').value=nf.conta||'';
  document.getElementById('nfAno').value=nf.ano||'2026';
  document.getElementById('tabManual').classList.add('active');
  document.getElementById('tabXML').classList.remove('active');
  document.querySelectorAll('#modalNF .tab').forEach((t,i)=>t.classList.toggle('active',i===0));
  gerarCamposParcelas();
  setTimeout(()=>{
    nf.parcelas.forEach((p,i)=>{
      const idx=i+1;
      const ve=document.getElementById(`parc${idx}_venc`);
      const va=document.getElementById(`parc${idx}_val`);
      const si=document.getElementById(`parc${idx}_sit`);
      if(ve) ve.value=p.vencimento||'';
      if(va) va.value=String(p.valor||'');
      if(si) si.value=p.situacao||'A VENCER';
    });
  },50);
  openModal('modalNF');
}
async function excluirNF(id) {
  if(!confirm('Excluir esta NF? Esta ação remove a linha da planilha e não pode ser desfeita.')) return;
  const nf = notas.find(x=>x.id===id);
  const res = await apiCall('excluirNota', { numero: nf.numero, ano: nf.ano });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }
  await reloadNotas();
  renderNotas();
  if (currentUser.role === 'admin') renderDashboard();
}

// ====================================================
// APROVAÇÃO
// ====================================================
function renderAprovacao() {
  const pendentes=notas.filter(n=>(n.validacao||'').indexOf('AGUARDANDO')!==-1);
  const div=document.getElementById('aprovacaoList');
  if(!pendentes.length){
    div.innerHTML=`<div class="empty"><div class="eicon">✅</div><div>Nenhuma NF aguardando aprovação</div></div>`;
    return;
  }
  div.innerHTML=pendentes.map(nf=>{
    const st=calcStatus(nf);
    const parcChips=nf.parcelas.map((p,i)=>`<span class="parcela-chip">${i+1}. ${fmtDate(p.vencimento)} · ${fmtMoney(parseMoney(p.valor))}</span>`).join('');
    const canApr=canApprove();
    return `<div class="approval-card">
      <div class="approval-info">
        <h3>NF #${nf.numero} — ${nf.fornecedor}</h3>
        <p>Conta: ${nf.conta} · Ano: ${nf.ano}</p>
        <div class="approval-meta">
          <span>📅 ${fmtDate(nf.dataFat)}</span>
          <span>💰 ${fmtMoney(parseMoney(nf.valorTotal))}</span>
          <span>${nf.numParcelas}x parcelas</span>
        </div>
        <div class="parcela-summary">${parcChips}</div>
      </div>
      <div class="approval-actions">
        <button class="btn btn-ghost btn-sm" onclick="verDetalheNF('${nf.id}')">👁 Ver</button>
        ${canApr?`<button class="btn btn-success btn-sm" onclick="aprovarNF('${nf.id}')">✅ Aprovar</button>
        <button class="btn btn-danger btn-sm" onclick="abrirRecusar('${nf.id}')">❌ Recusar</button>`:`<span style="font-size:12px;color:var(--muted)">Sem permissão</span>`}
      </div>
    </div>`;
  }).join('');
}
async function aprovarNF(id) {
  if(!canApprove()){alert('Sem permissão.');return;}
  const nf=notas.find(x=>x.id===id);
  const res = await apiCall('aprovarNota', { numero: nf.numero, ano: nf.ano });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }
  await reloadNotas();
  renderAprovacao();renderNotas();
  if (currentUser.role === 'admin') renderDashboard();
}
function abrirRecusar(id) {
  if(!canApprove()){alert('Sem permissão.');return;}
  recusaTarget=id;
  document.getElementById('recusaMotivo').value='';
  openModal('modalRecusar');
}
async function confirmarRecusa() {
  const motivo=document.getElementById('recusaMotivo').value.trim();
  if(!motivo){alert('Informe o motivo.');return;}
  const nf=notas.find(x=>x.id===recusaTarget);
  const res = await apiCall('recusarNota', { numero: nf.numero, ano: nf.ano, motivo });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }
  await reloadNotas();
  closeModal('modalRecusar');
  renderAprovacao();renderNotas();
  if (currentUser.role === 'admin') renderDashboard();
}

// ====================================================
// FORNECEDORES
// ====================================================
let filtFornLocal=[];
function renderFornecedores() {
  filtFornLocal=[...fornecedores].sort((a,b)=>a.nome.localeCompare(b.nome));
  renderFornecedoresTable();
}
function filterFornecedores() {
  const q=(document.getElementById('searchForn').value||'').toLowerCase();
  filtFornLocal=fornecedores.filter(f=>f.nome.toLowerCase().includes(q)||f.cnpj.includes(q));
  renderFornecedoresTable();
}
function renderFornecedoresTable() {
  const tb=document.getElementById('tbodyForn');
  if(!filtFornLocal.length){tb.innerHTML=`<tr><td colspan="7" class="empty">Nenhum fornecedor</td></tr>`;return;}
  tb.innerHTML=filtFornLocal.map((f,i)=>`<tr>
    <td style="color:var(--muted)">${i+1}</td>
    <td><strong>${f.nome}</strong></td>
    <td style="font-family:monospace;font-size:12px">${f.cnpj}</td>
    <td>${f.juros||'—'}</td>
    <td>${f.multa||'—'}</td>
    <td>${f.protesto||'—'}</td>
    <td>
      <div class="tbl-actions">
        <button class="btn btn-ghost btn-sm" onclick="editarFornecedor(${f.id})">✏️</button>
        <button class="btn btn-danger btn-sm" onclick="excluirFornecedor(${f.id})">🗑</button>
      </div>
    </td>
  </tr>`).join('');
}
async function salvarFornecedor() {
  const nome=document.getElementById('fNome').value.trim();
  const cnpj=document.getElementById('fCNPJ').value.trim();
  if(!nome||!cnpj){alert('Preencha nome e CNPJ.');return;}
  const existing=fornecedores.find(f=>f.cnpj===cnpj&&f.nome===nome);
  if(existing){alert('Fornecedor já cadastrado.');return;}

  const forn = {nome,cnpj,juros:parseFloat(document.getElementById('fJuros').value)||0,multa:parseFloat(document.getElementById('fMulta').value)||0,protesto:parseInt(document.getElementById('fProtesto').value)||0};
  const res = await apiCall('salvarFornecedor', { fornecedor: JSON.stringify(forn) });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }

  const r = await apiCall('listarFornecedores', {});
  if(r.ok) fornecedores = r.fornecedores;
  populateSelects();renderFornecedores();closeModal('modalFornecedor');
  ['fNome','fCNPJ','fJuros','fMulta','fProtesto'].forEach(i=>{const el=document.getElementById(i);if(el)el.value='';});
}
function editarFornecedor(id){
  const f=fornecedores.find(x=>x.id===id);
  if(!f)return;
  document.getElementById('fNome').value=f.nome;
  document.getElementById('fCNPJ').value=f.cnpj;
  document.getElementById('fJuros').value=f.juros||'';
  document.getElementById('fMulta').value=f.multa||'';
  document.getElementById('fProtesto').value=f.protesto||'';
  openModal('modalFornecedor');
  const btnSave = document.querySelector('#modalFornecedor .btn-primary');
  const origSave=btnSave.onclick;
  btnSave.onclick=async function(){
    const forn = {
      nome: document.getElementById('fNome').value.trim(),
      cnpj: document.getElementById('fCNPJ').value.trim(),
      juros: parseFloat(document.getElementById('fJuros').value)||0,
      multa: parseFloat(document.getElementById('fMulta').value)||0,
      protesto: parseInt(document.getElementById('fProtesto').value)||0
    };
    const res = await apiCall('salvarFornecedor', { fornecedor: JSON.stringify(forn) });
    if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }
    const r = await apiCall('listarFornecedores', {});
    if(r.ok) fornecedores = r.fornecedores;
    populateSelects();renderFornecedores();closeModal('modalFornecedor');
    btnSave.onclick=origSave;
  };
}
async function excluirFornecedor(id){
  if(!confirm('Excluir fornecedor?'))return;
  const f=fornecedores.find(x=>x.id===id);
  const res = await apiCall('excluirFornecedor', { cnpj: f.cnpj });
  if(!res.ok){ alert('Erro: '+(res.erro||'')); return; }
  const r = await apiCall('listarFornecedores', {});
  if(r.ok) fornecedores = r.fornecedores;
  renderFornecedores();populateSelects();
}

// ====================================================
// DASHBOARD
// ====================================================
let charts={};
function renderDashboard() {
  if (!currentUser || currentUser.role !== 'admin') return; // dashboard restrito a administradores
  const yr=document.getElementById('dashYear').value;
  const filtered=yr==='all'?notas:notas.filter(n=>n.ano===yr);
  const now=new Date();now.setHours(0,0,0,0);
  const em15=new Date(now);em15.setDate(em15.getDate()+15);

  let totalValor=0,totalPago=0,totalVencido=0,totalProximo=0,countVenc=0,countPago=0,countAvencer=0,countProximo=0;
  const parcelaVencidas=[],parcelaProximas=[];

  filtered.forEach(nf=>{
    totalValor+=parseMoney(nf.valorTotal);
    const st=calcStatus(nf);
    if(st==='PAGO') countPago++;
    else if(st==='VENCIDA') countVenc++;
    else if(st==='PRÓXIMO VENCIMENTO') countProximo++;
    else countAvencer++;
    (nf.parcelas||[]).forEach((p,i)=>{
      const v=new Date(p.vencimento);v.setHours(0,0,0,0);
      const pago=String(p.situacao||'').toUpperCase()==='PAGO';
      const val=parseMoney(p.valor);
      if(pago) totalPago+=val;
      else if(v<now){totalVencido+=val;parcelaVencidas.push({nf,p,i});}
      else if(v>=now&&v<=em15){totalProximo+=val;parcelaProximas.push({nf,p,i});}
    });
  });

  const kpis=[
    {label:'Total de NFs',value:filtered.length,sub:`${yr==='all'?'Todos os anos':yr}`,cls:''},
    {label:'Valor Total',value:fmtMoney(totalValor),sub:'emitido',cls:''},
    {label:'Pago',value:fmtMoney(totalPago),sub:`${countPago} notas quitadas`,cls:'green'},
    {label:'Vencidas',value:fmtMoney(totalVencido),sub:`${parcelaVencidas.length} parcelas`,cls:'red'},
    {label:'Próx. Vencimento',value:fmtMoney(totalProximo),sub:`${parcelaProximas.length} parcelas (≤15 dias)`,cls:'yellow'},
    {label:'Aguardando Aprovação',value:filtered.filter(n=>n.validacao==='AGUARDANDO VALIDAÇÃO').length,sub:'NFs pendentes',cls:'orange'},
  ];
  document.getElementById('kpiGrid').innerHTML=kpis.map(k=>`
    <div class="kpi ${k.cls}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');

  // Charts
  const statusCounts={PAGO:countPago,VENCIDA:countVenc,'PRÓXIMO VENCIMENTO':countProximo,'A VENCER':countAvencer};
  renderChart('chartStatus','doughnut',Object.keys(statusCounts),Object.values(statusCounts),['#10b981','#ef4444','#f59e0b','#3b82f6']);

  const contasMap={};
  filtered.forEach(n=>{const c=n.conta||'OUTROS';contasMap[c]=(contasMap[c]||0)+parseMoney(n.valorTotal);});
  const contasSorted=Object.entries(contasMap).sort((a,b)=>b[1]-a[1]).slice(0,8);
  renderChart('chartContas','bar',contasSorted.map(x=>x[0]),contasSorted.map(x=>x[1]),['#6366f1'],true);

  const mesesMap={};
  filtered.forEach(nf=>(nf.parcelas||[]).forEach(p=>{
    if(!p.vencimento) return;
    const d=new Date(p.vencimento);
    const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    mesesMap[key]=(mesesMap[key]||0)+parseMoney(p.valor);
  }));
  const mesesSorted=Object.entries(mesesMap).sort((a,b)=>a[0].localeCompare(b[0])).slice(-12);
  const mesesLabels=mesesSorted.map(([k])=>{const[y,m]=k.split('-');return `${['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][parseInt(m)-1]}/${y.slice(2)}`});
  renderChart('chartMeses','line',mesesLabels,mesesSorted.map(x=>x[1]),['#3b82f6']);

  const fornMap={};
  filtered.forEach(n=>{const f=n.fornecedor||'—';fornMap[f]=(fornMap[f]||0)+parseMoney(n.valorTotal);});
  const fornTop=Object.entries(fornMap).sort((a,b)=>b[1]-a[1]).slice(0,10);
  renderChart('chartForn','bar',fornTop.map(x=>x[0]),fornTop.map(x=>x[1]),['#f97316'],true,true);

  // Alertas
  const alertSec=document.getElementById('alertasSection');
  let alertHTML='';
  if(parcelaVencidas.length){
    alertHTML+=`<div class="alert-section">
      <div class="alert-header red">🔴 VENCIDAS — ${parcelaVencidas.length} parcelas · ${fmtMoney(totalVencido)}</div>
      <div class="alert-table" style="overflow-x:auto">
      <table class="tbl"><thead><tr><th>NF</th><th>Fornecedor</th><th>Parcela</th><th>Vencimento</th><th>Valor</th><th>Ação</th></tr></thead>
      <tbody>${parcelaVencidas.slice(0,20).map(({nf,p,i})=>`<tr>
        <td>#${nf.numero}</td><td>${nf.fornecedor}</td><td>${i+1}</td>
        <td style="color:#fca5a5">${fmtDate(p.vencimento)}</td>
        <td>${fmtMoney(parseMoney(p.valor))}</td>
        <td><button class="btn btn-success btn-sm" onclick="abrirBaixar('${nf.id}',${i})">💰 Baixar</button></td>
      </tr>`).join('')}</tbody></table></div></div>`;
  }
  if(parcelaProximas.length){
    alertHTML+=`<div class="alert-section">
      <div class="alert-header yellow">🟡 PRÓXIMO VENCIMENTO — ${parcelaProximas.length} parcelas · ${fmtMoney(totalProximo)}</div>
      <div class="alert-table" style="overflow-x:auto">
      <table class="tbl"><thead><tr><th>NF</th><th>Fornecedor</th><th>Parcela</th><th>Vencimento</th><th>Valor</th><th>Ação</th></tr></thead>
      <tbody>${parcelaProximas.map(({nf,p,i})=>`<tr>
        <td>#${nf.numero}</td><td>${nf.fornecedor}</td><td>${i+1}</td>
        <td style="color:#fcd34d">${fmtDate(p.vencimento)}</td>
        <td>${fmtMoney(parseMoney(p.valor))}</td>
        <td><button class="btn btn-success btn-sm" onclick="abrirBaixar('${nf.id}',${i})">💰 Baixar</button></td>
      </tr>`).join('')}</tbody></table></div></div>`;
  }
  alertSec.innerHTML=alertHTML;
}

function renderChart(id,type,labels,data,colors,yMoney=false,horizontal=false){
  if(charts[id]) charts[id].destroy();
  const ctx=document.getElementById(id);
  if(!ctx) return;
  const isDoughnut=type==='doughnut';
  charts[id]=new Chart(ctx,{
    type:horizontal?'bar':type,
    data:{labels,datasets:[{data,backgroundColor:isDoughnut?colors:colors[0],borderColor:isDoughnut?'transparent':colors[0],fill:type==='line'?false:undefined,tension:.4,borderWidth:type==='line'?2:0,pointRadius:type==='line'?3:0}]},
    options:{
      indexAxis:horizontal?'y':'x',
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:isDoughnut,position:'right',labels:{color:'#94a3b8',font:{size:11},padding:12,boxWidth:12}},tooltip:{callbacks:{label:ctx=>yMoney||(!isDoughnut&&type!=='line')?fmtMoney(ctx.raw):ctx.raw}}},
      scales:isDoughnut?{}:{x:{ticks:{color:'#64748b',font:{size:10},maxRotation:45},grid:{color:'#1e293b'}},y:{ticks:{color:'#64748b',font:{size:10},callback:v=>yMoney?fmtMoney(v):v},grid:{color:'#1e293b'}}}
    }
  });
}

// ====================================================
// EXPORT
// ====================================================
function exportXLS() {
  const data=filteredNotas.length?filteredNotas:notas;
  const rows=data.map(n=>{
    const st=calcStatus(n);
    const base={Nº_NF:n.numero,Data_Faturamento:fmtDate(n.dataFat),Fornecedor:n.fornecedor,CNPJ:n.cnpj,Valor_Total:parseMoney(n.valorTotal),Num_Parcelas:n.numParcelas,Conta:n.conta,Ano:n.ano,Validacao:n.validacao,Status:st};
    (n.parcelas||[]).forEach((p,i)=>{base[`P${i+1}_Venc`]=fmtDate(p.vencimento);base[`P${i+1}_Valor`]=parseMoney(p.valor);base[`P${i+1}_Sit`]=p.situacao;base[`P${i+1}_DataPag`]=p.dataPagamento?fmtDate(p.dataPagamento):'';base[`P${i+1}_ValPago`]=p.valorPago?parseMoney(p.valorPago):'';});
    return base;
  });
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Notas Fiscais');
  XLSX.writeFile(wb,`NFs_${new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')}.xlsx`);
}

function exportPDF() {
  const data=filteredNotas.length?filteredNotas:notas;
  const html=buildPDFHTML('Relatório de Notas Fiscais',data);
  const w=window.open('','_blank');
  w.document.write(html);w.document.close();w.print();
}

function buildPDFHTML(titulo,data) {
  const rows=data.map(n=>{
    const st=calcStatus(n);
    return `<tr><td>${n.numero}</td><td>${fmtDate(n.dataFat)}</td><td>${n.fornecedor}</td><td>${fmtMoney(parseMoney(n.valorTotal))}</td><td>${n.conta}</td><td>${n.validacao}</td><td>${st}</td></tr>`;
  }).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${titulo}</title>
  <style>body{font-family:Arial;font-size:12px;color:#000}h1{font-size:16px;margin-bottom:8px}p{font-size:11px;color:#666;margin-bottom:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}th{background:#f0f0f0;font-weight:bold}tr:nth-child(even){background:#f9f9f9}@media print{body{margin:0}}</style>
  </head><body><h1>${titulo}</h1><p>Gerado em ${new Date().toLocaleString('pt-BR')} | Total: ${data.length} notas</p>
  <table><thead><tr><th>Nº NF</th><th>Data Fat.</th><th>Fornecedor</th><th>Valor Total</th><th>Conta</th><th>Validação</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></body></html>`;
}

// RELATÓRIOS
function getRelatorioData(tipo) {
  const now=new Date();now.setHours(0,0,0,0);
  const em15=new Date(now);em15.setDate(em15.getDate()+15);
  if(tipo==='vencidas'){
    const rows=[];
    notas.forEach(nf=>(nf.parcelas||[]).forEach((p,i)=>{
      const v=new Date(p.vencimento);v.setHours(0,0,0,0);
      if(v<now&&p.situacao!=='PAGO') rows.push({nf,p,idx:i+1});
    }));
    return rows.sort((a,b)=>new Date(a.p.vencimento)-new Date(b.p.vencimento));
  }
  if(tipo==='proximos'){
    const rows=[];
    notas.forEach(nf=>(nf.parcelas||[]).forEach((p,i)=>{
      const v=new Date(p.vencimento);v.setHours(0,0,0,0);
      if(v>=now&&v<=em15&&p.situacao!=='PAGO') rows.push({nf,p,idx:i+1});
    }));
    return rows.sort((a,b)=>new Date(a.p.vencimento)-new Date(b.p.vencimento));
  }
  if(tipo==='periodo'){
    const de=document.getElementById('relDe').value;
    const ate=document.getElementById('relAte').value;
    return notas.filter(n=>{
      if(!n.dataFat) return false;
      const d=n.dataFat.slice(0,10);
      if(de&&d<de) return false;
      if(ate&&d>ate) return false;
      return true;
    });
  }
  if(tipo==='fornecedor'){
    const fn=document.getElementById('relFornecedor').value;
    return notas.filter(n=>!fn||n.fornecedor===fn);
  }
  return [];
}

function gerarRelatorioXLS(tipo) {
  const data=getRelatorioData(tipo);
  if(!data.length){alert('Nenhum dado para exportar.');return;}
  let rows;
  if(tipo==='vencidas'||tipo==='proximos'){
    rows=data.map(({nf,p,idx})=>({Nº_NF:nf.numero,Fornecedor:nf.fornecedor,CNPJ:nf.cnpj,Parcela:idx,Vencimento:fmtDate(p.vencimento),Valor:parseMoney(p.valor),Situacao:p.situacao,Conta:nf.conta}));
  } else {
    rows=data.map(n=>({Nº_NF:n.numero,Data_Fat:fmtDate(n.dataFat),Fornecedor:n.fornecedor,Valor_Total:parseMoney(n.valorTotal),Conta:n.conta,Status:calcStatus(n),Validacao:n.validacao}));
  }
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Relatório');
  XLSX.writeFile(wb,`Relatorio_${tipo}_${new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')}.xlsx`);
}

function gerarRelatorioPDF(tipo) {
  const data=getRelatorioData(tipo);
  if(!data.length){alert('Nenhum dado para exportar.');return;}
  const titulos={vencidas:'Parcelas Vencidas',proximos:'Próximo Vencimento (15 dias)',periodo:'Por Período',fornecedor:'Por Fornecedor'};
  let rows;
  if(tipo==='vencidas'||tipo==='proximos'){
    rows=data.map(({nf,p,idx})=>`<tr><td>${nf.numero}</td><td>${nf.fornecedor}</td><td>${idx}</td><td>${fmtDate(p.vencimento)}</td><td>${fmtMoney(parseMoney(p.valor))}</td><td>${nf.conta}</td></tr>`).join('');
    const h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${titulos[tipo]}</title><style>body{font-family:Arial;font-size:12px}h1{font-size:16px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px}th{background:#f0f0f0}@media print{body{margin:0}}</style></head><body><h1>${titulos[tipo]}</h1><p>Gerado: ${new Date().toLocaleString('pt-BR')} | ${data.length} parcelas</p><table><thead><tr><th>NF</th><th>Fornecedor</th><th>Parc.</th><th>Vencimento</th><th>Valor</th><th>Conta</th></tr></thead><tbody>${rows}</tbody></table></body></html>`;
    const w=window.open('','_blank');w.document.write(h);w.document.close();w.print();
  } else {
    const html=buildPDFHTML(titulos[tipo],data);
    const w=window.open('','_blank');w.document.write(html);w.document.close();w.print();
  }
}

function gerarRelatorio(tipo) {
  const data=getRelatorioData(tipo);
  const titulos={vencidas:'🔴 Parcelas Vencidas',proximos:'🟡 Próximo Vencimento',periodo:'📅 Por Período',fornecedor:'🏢 Por Fornecedor'};
  const res=document.getElementById('relResultado');
  if(!data.length){res.innerHTML=`<div class="notice notice-info">Nenhum dado encontrado para este filtro.</div>`;return;}
  let html=`<div class="table-wrap"><div class="table-toolbar" style="justify-content:space-between"><div style="font-weight:700">${titulos[tipo]} — ${data.length} registros</div>
  <div style="display:flex;gap:8px"><button class="btn btn-ghost btn-sm" onclick="gerarRelatorioXLS('${tipo}')">📥 XLS</button><button class="btn btn-ghost btn-sm" onclick="gerarRelatorioPDF('${tipo}')">📄 PDF</button></div></div>
  <div style="overflow-x:auto"><table class="tbl"><thead>`;
  if(tipo==='vencidas'||tipo==='proximos'){
    html+=`<tr><th>Nº NF</th><th>Fornecedor</th><th>Parcela</th><th>Vencimento</th><th>Valor</th><th>Conta</th><th>Status</th></tr></thead><tbody>`;
    html+=data.map(({nf,p,idx})=>`<tr><td>#${nf.numero}</td><td>${nf.fornecedor}</td><td>${idx}</td><td>${fmtDate(p.vencimento)}</td><td>${fmtMoney(parseMoney(p.valor))}</td><td>${nf.conta||'—'}</td><td>${statusBadge(p.situacao||'A VENCER')}</td></tr>`).join('');
  } else {
    html+=`<tr><th>Nº NF</th><th>Data Fat.</th><th>Fornecedor</th><th>Valor Total</th><th>Conta</th><th>Parcelas</th><th>Validação</th><th>Status</th></tr></thead><tbody>`;
    html+=data.map(n=>`<tr><td>#${n.numero}</td><td>${fmtDate(n.dataFat)}</td><td>${n.fornecedor}</td><td>${fmtMoney(parseMoney(n.valorTotal))}</td><td>${n.conta||'—'}</td><td>${n.numParcelas||1}x</td><td>${validBadge(n.validacao)}</td><td>${statusBadge(calcStatus(n))}</td></tr>`).join('');
  }
  html+=`</tbody></table></div></div>`;
  res.innerHTML=html;
}

// ====================================================
// INIT
// ====================================================
window.onload=()=>{
  // default date inputs
  document.getElementById('relDe').value=new Date(new Date().getFullYear(),0,1).toISOString().slice(0,10);
  document.getElementById('relAte').value=today();

  if (GAS_URL) {
    document.getElementById('configScreen').style.display='none';
    document.getElementById('loginScreen').style.display='flex';
    document.getElementById('loginUser').focus();
  } else {
    document.getElementById('configScreen').style.display='flex';
    document.getElementById('loginScreen').style.display='none';
  }
};
</script>
</body>
</html>
