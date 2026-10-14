# VN-Dashboard
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IAELO - Vietnamese Students Directory & Compliance Tracker</title>

<!-- Google Fonts -->
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<!-- SheetJS (XLSX) Library for Data Import/Export -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

<style>
:root {
  --primary: #0f5132;
  --primary-light: #d8ebd9;
  --primary-hover: #083321;
  --accent: #2e7d32;
  --accent-light: #c8e6c9;
  --bg-light: #e4f0e8;
  --white: #ffffff;
  --text-dark: #112217;
  --text-gray: #3d5846;
  --danger: #dc3545;
  --danger-light: #f8d7da;
  --warning: #d97706;
  --warning-light: #fef3c7;
  --border: #b8dfc1;
  --shadow: 0 6px 24px rgba(15, 81, 50, 0.09);
  --radius: 16px;
  --btn-radius: 10px;
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
  font-family: 'Plus Jakarta Sans', sans-serif;
}

body {
  background-color: var(--bg-light);
  color: var(--text-dark);
  min-height: 100vh;
  font-weight: 400;
}

/* LOGIN / SIGN UP OVERLAY */
.login-overlay {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: linear-gradient(135deg, rgba(15, 81, 50, 0.94), rgba(8, 51, 33, 0.96));
  backdrop-filter: blur(8px);
  display: flex;
  justify-content: center;
  align-items: center;
  z-index: 9999;
}

.login-card {
  background: var(--white);
  padding: 35px 40px;
  border-radius: var(--radius);
  width: 100%;
  max-width: 440px;
  border: 1px solid var(--border);
  box-shadow: 0 15px 35px rgba(0,0,0,0.3);
  text-align: center;
}

.login-card h2 {
  color: var(--primary);
  font-size: 22px;
  font-weight: 800;
  margin-bottom: 6px;
}

.login-card p.subheading {
  color: var(--text-gray);
  font-size: 13px;
  margin-bottom: 20px;
}

.auth-msg {
  padding: 10px;
  border-radius: 8px;
  font-size: 12px;
  font-weight: 600;
  margin-bottom: 16px;
  display: none;
}

.auth-msg.error {
  background: var(--danger-light);
  color: var(--danger);
  border: 1px solid #f5c6cb;
}

.auth-msg.success {
  background: var(--primary-light);
  color: var(--primary);
  border: 1px solid var(--border);
}

.auth-toggle-bottom {
  margin-top: 18px;
  font-size: 13px;
  color: var(--text-gray);
}

.auth-toggle-link {
  color: var(--primary);
  font-weight: 700;
  cursor: pointer;
  text-decoration: underline;
}

/* BRANDING & LOGO HEADER BAR */
.brand-bar {
  background: linear-gradient(90deg, #ffffff, #dcf2e1);
  padding: 16px 40px;
  border-bottom: 2px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 20px;
}

.brand-logo-container {
  display: flex;
  align-items: center;
  gap: 16px;
}

.brand-logo-img {
  width: 65px;
  height: 65px;
  border-radius: 50%;
  object-fit: cover;
  border: 3px solid var(--primary);
  background: #ffffff;
  padding: 2px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.1);
  cursor: pointer;
  transition: transform 0.2s;
}

.brand-logo-img:hover {
  transform: scale(1.05);
}

.brand-info h2 {
  font-size: 20px;
  color: var(--primary);
  font-weight: 800;
  letter-spacing: -0.3px;
}

.brand-info p {
  font-size: 13px;
  color: var(--text-gray);
  font-weight: 600;
}

/* DASHBOARD HEADER STYLES */
header {
  background: linear-gradient(135deg, var(--primary), var(--primary-hover));
  padding: 18px 40px;
  border-bottom: 2px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: center;
  box-shadow: 0 4px 15px rgba(0,0,0,0.1);
  position: relative;
}

.header-left-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.notification-bell-wrapper {
  position: relative;
}

.bell-btn {
  background: rgba(255, 255, 255, 0.15);
  border: 1px solid rgba(255, 255, 255, 0.3);
  color: #ffffff;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s;
  position: relative;
  padding: 0;
}

.bell-icon-svg {
  width: 15px;
  height: 15px;
  fill: #ffffff;
}

.bell-btn:hover {
  background: rgba(255, 255, 255, 0.3);
  border-color: rgba(255, 255, 255, 0.6);
}

.bell-badge {
  position: absolute;
  top: -3px;
  right: -3px;
  background: var(--danger);
  color: #ffffff;
  font-size: 9px;
  font-weight: 800;
  border-radius: 8px;
  padding: 1px 4px;
  border: 1.5px solid var(--primary);
  display: none;
  line-height: 1;
}

/* NOTIFICATION DROPDOWN PANEL */
.notif-dropdown {
  display: none;
  position: absolute;
  top: 42px;
  left: 0;
  width: 400px;
  background: var(--white);
  border: 2px solid var(--border);
  border-radius: var(--radius);
  box-shadow: 0 10px 30px rgba(0,0,0,0.25);
  z-index: 1000;
  padding: 18px;
}

.notif-dropdown h4 {
  font-size: 14px;
  color: var(--primary);
  font-weight: 800;
  margin-bottom: 12px;
  border-bottom: 1.5px solid var(--primary-light);
  padding-bottom: 6px;
}

.header-title h1 {
  color: #ffffff;
  font-size: 20px;
  font-weight: 800;
  letter-spacing: 0.2px;
}

.header-title p {
  color: var(--accent-light);
  font-size: 12px;
  margin-top: 2px;
  font-weight: 500;
}

.header-right {
  display: flex;
  align-items: center;
  gap: 12px;
}

.user-badge {
  background: rgba(255, 255, 255, 0.2);
  color: #ffffff;
  padding: 8px 16px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 700;
  border: 1px solid rgba(255, 255, 255, 0.4);
}

.btn-logout {
  background: #ffffff;
  color: var(--danger);
  border: 1px solid #ffffff;
  padding: 8px 16px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.2s;
}

.btn-logout:hover {
  background: var(--danger);
  color: var(--white);
  border-color: var(--danger);
}

/* MAIN DASHBOARD LAYOUT */
.main-content {
  max-width: 1400px;
  margin: 30px auto;
  padding: 0 20px;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 20px;
  margin-bottom: 24px;
}

.stat-card {
  background: var(--white);
  padding: 22px 24px;
  border-radius: var(--radius);
  border: 1px solid var(--border);
  box-shadow: var(--shadow);
  position: relative;
  overflow: hidden;
  cursor: pointer;
  transition: transform 0.2s, box-shadow 0.2s;
}

.stat-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 28px rgba(15, 81, 50, 0.14);
}

.stat-card::before {
  content: "";
  position: absolute;
  top: 0;
  left: 0;
  width: 6px;
  height: 100%;
  background: var(--primary);
}

.stat-card.alert-card::before {
  background: var(--danger);
}

.stat-card.warning-card::before {
  background: var(--warning);
}

.stat-card h3 {
  font-size: 12px;
  color: var(--text-gray);
  margin-bottom: 8px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.stat-card .number {
  font-size: 34px;
  font-weight: 800;
  color: var(--primary);
}

.stat-card p.card-subtext {
  font-size: 11px;
  color: var(--text-gray);
  margin-top: 4px;
  font-weight: 600;
}

/* TOOLBAR */
.toolbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
  gap: 12px;
  flex-wrap: wrap;
}

.search-box {
  flex: 1;
  min-width: 280px;
  position: relative;
}

.search-box input {
  width: 100%;
  padding: 12px 16px 12px 40px;
  border: 2px solid var(--border);
  border-radius: var(--btn-radius);
  background: var(--white);
  font-size: 13px;
  font-weight: 600;
  outline: none;
  color: var(--text-dark);
  transition: all 0.2s;
}

.search-box input:focus {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px rgba(15, 81, 50, 0.15);
}

.search-box::before {
  content: "🔍";
  position: absolute;
  left: 14px;
  top: 12px;
  font-size: 14px;
}

.btn {
  background-color: var(--primary);
  color: var(--white);
  border: none;
  padding: 12px 20px;
  border-radius: var(--btn-radius);
  font-size: 14px;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.2s;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  box-shadow: 0 3px 10px rgba(15, 81, 50, 0.2);
}

.btn:hover {
  background-color: var(--primary-hover);
  transform: translateY(-1px);
}

.btn-secondary {
  background-color: var(--primary-light);
  color: var(--primary);
  border: 1.5px solid var(--border);
  box-shadow: none;
}

.btn-secondary:hover {
  background-color: var(--accent-light);
  color: var(--primary-hover);
}

.btn-danger {
  background-color: var(--danger);
  color: var(--white);
  box-shadow: 0 3px 10px rgba(220, 53, 69, 0.2);
}

.btn-danger:hover {
  background-color: #b02a37;
}

.student-list-container {
  background: var(--white);
  border-radius: var(--radius);
  border: 1px solid var(--border);
  box-shadow: var(--shadow);
  overflow: hidden;
  padding: 12px;
}

.student-list-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 16px 20px;
  border: 1.5px solid #e2eee5;
  background: #ffffff;
  cursor: pointer;
  transition: all 0.2s ease-in-out;
  border-radius: var(--btn-radius);
  margin-bottom: 8px;
}

.student-list-item:hover {
  background: var(--primary-light);
  border-color: var(--border);
  transform: translateX(4px);
}

.student-list-item.active-student {
  background-color: #cce8d3 !important;
  border: 2px solid var(--primary) !important;
  box-shadow: 0 4px 12px rgba(15, 81, 50, 0.15);
}

.student-list-item.active-student .click-hint {
  background: var(--primary);
  color: var(--white);
  border-color: var(--primary);
}

.student-list-left {
  display: flex;
  align-items: center;
  gap: 16px;
}

.student-thumb {
  width: 50px;
  height: 50px;
  border-radius: 50%;
  object-fit: cover;
  border: 2px solid var(--primary);
  background: var(--white);
  box-shadow: 0 2px 6px rgba(0,0,0,0.1);
}

.student-basic-info h4 {
  color: var(--primary);
  font-size: 16px;
  font-weight: 800;
  margin-bottom: 2px;
}

.student-basic-info p {
  color: var(--text-gray);
  font-size: 13px;
  font-weight: 500;
}

.click-hint {
  color: var(--primary);
  font-size: 12px;
  font-weight: 700;
  background: var(--primary-light);
  padding: 8px 14px;
  border-radius: 20px;
  border: 1px solid var(--border);
  transition: all 0.2s;
}

.request-card {
  display: flex;
  flex-direction: column;
  padding: 10px 12px;
  background: var(--primary-light);
  border-radius: 8px;
  border: 1px solid var(--border);
  margin-bottom: 8px;
  gap: 8px;
}

.request-info p {
  font-size: 12px;
  font-weight: 700;
  color: var(--text-dark);
}

.request-info span {
  font-size: 11px;
  color: var(--text-gray);
  font-weight: 500;
}

.request-actions {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
}

/* MODAL STYLES */
.modal {
  display: none;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(15, 40, 25, 0.75);
  backdrop-filter: blur(4px);
  justify-content: center;
  align-items: center;
  z-index: 1000;
  padding: 20px;
}

.modal-content {
  background: var(--white);
  padding: 30px;
  border-radius: var(--radius);
  width: 100%;
  max-width: 620px;
  max-height: 90vh;
  overflow-y: auto;
  border: 2px solid var(--border);
  box-shadow: 0 10px 40px rgba(0,0,0,0.25);
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
  border-bottom: 2px solid var(--primary-light);
  padding-bottom: 12px;
}

.modal-header h3 {
  color: var(--primary);
  font-size: 20px;
  font-weight: 800;
}

.close-btn {
  background: none;
  border: none;
  font-size: 24px;
  cursor: pointer;
  color: var(--text-gray);
  transition: color 0.2s;
}

.close-btn:hover {
  color: var(--danger);
}

.profile-view-header {
  display: flex;
  gap: 20px;
  align-items: center;
  margin-bottom: 20px;
  background: var(--primary-light);
  padding: 18px;
  border-radius: 12px;
  border: 1px solid var(--border);
}

.profile-view-img {
  width: 85px;
  height: 85px;
  border-radius: 50%;
  object-fit: cover;
  border: 3px solid var(--primary);
  box-shadow: 0 4px 10px rgba(0,0,0,0.1);
}

.info-row {
  display: flex;
  justify-content: space-between;
  margin-bottom: 10px;
  font-size: 13px;
  border-bottom: 1px dashed var(--border);
  padding-bottom: 6px;
}

.info-row span:first-child {
  color: var(--text-gray);
  font-weight: 600;
}

.info-row span:last-child {
  color: var(--text-dark);
  font-weight: 700;
  text-align: right;
}

.compliance-box {
  background: #eaf5ec;
  padding: 16px;
  border-radius: 12px;
  margin-top: 18px;
  border: 1.5px solid var(--border);
}

.compliance-box h5 {
  font-size: 12px;
  color: var(--primary);
  margin-bottom: 10px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  font-weight: 800;
}

.form-group {
  margin-bottom: 14px;
  text-align: left;
}

.form-group label {
  display: block;
  font-size: 12px;
  color: var(--text-gray);
  margin-bottom: 5px;
  font-weight: 700;
}

.form-control {
  width: 100%;
  padding: 10px 14px;
  border: 1.5px solid var(--border);
  border-radius: var(--btn-radius);
  font-size: 13px;
  font-weight: 600;
  color: var(--text-dark);
  outline: none;
  transition: all 0.2s;
  background: #ffffff;
}

.form-control:focus {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px rgba(15, 81, 50, 0.15);
}

.form-control[readonly] {
  background-color: var(--primary-light);
  color: var(--primary);
  cursor: not-allowed;
}

.form-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 15px;
}

.form-grid .form-group.full-width {
  grid-column: span 2;
}

.alert-badge {
  display: inline-block;
  padding: 6px 12px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 700;
  margin-top: 4px;
}
.alert-danger { background: var(--danger-light); color: var(--danger); border: 1px solid #f5c6cb; }
.alert-warning { background: var(--warning-light); color: var(--warning); border: 1px solid #fde68a; }
.alert-success { background: var(--primary-light); color: var(--primary); border: 1px solid var(--border); }
.alert-info { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }

.photo-preview-box {
  display: flex;
  align-items: center;
  gap: 15px;
  margin-top: 8px;
}

.photo-preview-img {
  width: 55px;
  height: 55px;
  border-radius: 50%;
  object-fit: cover;
  border: 2px solid var(--primary);
  display: none;
}
</style>
</head>
<body>

<!-- LOGIN / SIGN UP OVERLAY -->
<div id="loginOverlay" class="login-overlay">
  <div class="login-card">
    <h2 id="authHeading">IAELO Portal Access</h2>
    <p id="authSubheading" class="subheading">Enter your credentials or request access</p>
    
    <div id="authMessage" class="auth-msg"></div>

    <!-- LOGIN FORM -->
    <form id="loginForm" onsubmit="handleLogin(event)">
      <div class="form-group">
        <label>Email Address</label>
        <input type="email" id="loginEmail" class="form-control" placeholder="Enter your email address" required>
      </div>
      <div class="form-group">
        <label>Password</label>
        <input type="password" id="loginPassword" class="form-control" placeholder="Enter password" required>
      </div>
      <button type="submit" class="btn" style="width: 100%; margin-top: 10px;">Log In</button>
      <div class="auth-toggle-bottom">
        Need access? <span class="auth-toggle-link" onclick="toggleAuthMode('register')">Request Access</span>
      </div>
    </form>

    <!-- REGISTER FORM -->
    <form id="registerForm" onsubmit="handleRegister(event)" style="display: none;">
      <div class="form-group">
        <label>Full Name</label>
        <input type="text" id="regName" class="form-control" placeholder="Enter your full name" required>
      </div>
      <div class="form-group">
        <label>Email Address</label>
        <input type="email" id="regEmail" class="form-control" placeholder="Enter your email address" required>
      </div>
      <div class="form-group">
        <label>Password</label>
        <input type="password" id="regPassword" class="form-control" placeholder="Create a password (min 6 chars)" required minlength="6">
      </div>
      <div class="form-group">
        <label>Confirm Password</label>
        <input type="password" id="regConfirmPassword" class="form-control" placeholder="Confirm your password" required minlength="6">
      </div>
      <button type="submit" class="btn" style="width: 100%; margin-top: 10px;">Submit Access Request</button>
      <div class="auth-toggle-bottom">
        Already have an account? <span class="auth-toggle-link" onclick="toggleAuthMode('login')">Log In</span>
      </div>
    </form>
  </div>
</div>

<!-- MAIN DASHBOARD APPLICATION -->
<div id="app-container">
  
  <!-- BRAND & LOGO BAR -->
  <div class="brand-bar">
    <div class="brand-logo-container">
      <img id="appLogo" src="https://via.placeholder.com/65?text=LOGO" class="brand-logo-img" alt="University Logo" onclick="openLogoModal()" title="Click to view/manage logo">
      <div class="brand-info">
        <h2>Southern Luzon State University</h2>
        <p>International Affairs & External Linkages Office (IAELO)</p>
      </div>
    </div>
  </div>

  <!-- HEADER BANNER -->
  <header>
    <div class="header-left-actions">
      <!-- NOTIFICATION BELL & DROPDOWN FOR PENDING ACCESS REQUESTS -->
      <div class="notification-bell-wrapper">
        <button class="bell-btn" onclick="toggleNotifDropdown()" title="Notifications / Pending Requests">
          <svg class="bell-icon-svg" viewBox="0 0 24 24">
            <path d="M12 22c1.1 0 2-.9 2-2h-4c0 1.1.89 2 2 2zm6-6v-5c0-3.07-1.64-5.64-4.5-6.32V4c0-.83-.67-1.5-1.5-1.5s-1.5.67-1.5 1.5v.68C7.63 5.36 6 7.92 6 11v5l-2 2v1h16v-1l-2-2z"/>
          </svg>
          <span id="bellBadge" class="bell-badge">0</span>
        </button>

        <div id="notifDropdown" class="notif-dropdown">
          <h4>📥 Pending Access Requests</h4>
          <div id="pendingRequestsContainer">
            <p style="font-size: 12px; color: var(--text-gray);">No pending access requests.</p>
          </div>
        </div>
      </div>

      <div class="header-title">
        <h1>IAELO Compliance Dashboard</h1>
        <p>Vietnamese Students Tracking & Regulatory Portal</p>
      </div>
    </div>

    <div class="header-right">
      <div id="currentUserDisplay" class="user-badge">User: Loading...</div>
      <button class="btn-logout" onclick="handleLogout()">Logout</button>
    </div>
  </header>

  <div class="main-content">
    <!-- SUMMARY STATS -->
    <div class="stats-grid">
      <div class="stat-card" onclick="setQuickFilter('all')">
        <h3>Total Vietnamese Students</h3>
        <div class="number" id="totalStudentsCount">0</div>
        <p class="card-subtext">Click to show all student records</p>
      </div>
      <div class="stat-card alert-card" onclick="setQuickFilter('visa-alerts')">
        <h3>Visa Expiry Alerts (Next 60 Days)</h3>
        <div class="number" id="visaAlertsCount" style="color: var(--danger);">0</div>
        <p class="card-subtext">Click to filter upcoming visa expiries</p>
      </div>
      <div class="stat-card warning-card" onclick="setQuickFilter('threshold-alerts')">
        <h3>6-Month Entry Threshold Check</h3>
        <div class="number" id="thresholdAlertsCount" style="color: var(--warning);">0</div>
        <p class="card-subtext">Click to view 6-month window status</p>
      </div>
    </div>

    <!-- TOOLBAR -->
    <div class="toolbar">
      <div class="search-box">
        <input type="text" id="searchInput" placeholder="Search by Name, Course, Status, or Passport..." oninput="renderStudents()">
      </div>
      <button class="btn btn-secondary" onclick="exportToExcel()">📊 Export Excel</button>
      <button class="btn btn-secondary edit-allowed" onclick="document.getElementById('excelFileInput').click()">📥 Import Excel</button>
      <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" style="display: none;" onchange="importFromExcel(event)">
      <button class="btn edit-allowed" onclick="openAddModal()">+ Add New Student Record</button>
    </div>

    <!-- CLICKABLE STUDENT LIST CONTAINER -->
    <div class="student-list-container" id="studentListContainer"></div>
  </div>
</div>

<!-- LOGO DISPLAY & UPLOAD MODAL -->
<div id="logoModal" class="modal">
  <div class="modal-content" style="max-width: 420px; text-align: center;">
    <div class="modal-header">
      <h3>University Logo</h3>
      <button class="close-btn" onclick="closeLogoModal()">&times;</button>
    </div>
    <div style="padding: 10px;">
      <img id="modalLogoPreview" src="https://via.placeholder.com/180?text=LOGO" style="width: 180px; height: 180px; border-radius: 50%; object-fit: cover; border: 4px solid var(--primary); margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.15);" alt="Logo Full View">
      
      <div class="edit-allowed">
        <button class="btn" style="width: 100%;" onclick="document.getElementById('logoFileInput').click()">📷 Upload / Change Logo</button>
        <input type="file" id="logoFileInput" accept="image/*" style="display: none;" onchange="handleLogoUpload(event)">
      </div>
    </div>
  </div>
</div>

<!-- STUDENT DETAILS VIEW MODAL -->
<div id="viewModal" class="modal">
  <div class="modal-content">
    <div class="modal-header">
      <h3 id="viewModalTitle">Student Profile & Compliance Details</h3>
      <button class="close-btn" onclick="closeViewModal()">&times;</button>
    </div>
    <div id="viewModalBody"></div>
    <div style="display: flex; gap: 10px; margin-top: 20px;">
      <button class="btn edit-allowed" id="editFromViewBtn" style="flex: 1; background-color: var(--accent);">Edit Details</button>
      <button class="btn btn-danger edit-allowed" id="deleteFromViewBtn" style="flex: 1;">Delete Record</button>
      <button class="btn btn-secondary" onclick="closeViewModal()" style="flex: 1;">Close</button>
    </div>
  </div>
</div>

<!-- ADD/EDIT STUDENT MODAL -->
<div id="studentModal" class="modal">
  <div class="modal-content">
    <div class="modal-header">
      <h3 id="modalTitle">Add Vietnamese Student Record</h3>
      <button class="close-btn" onclick="closeModal()">&times;</button>
    </div>
    <form id="studentForm" onsubmit="saveStudent(event)">
      <input type="hidden" id="studentId">
      <input type="hidden" id="currentPhotoBase64">

      <div class="form-grid">
        <div class="form-group">
          <label>Nationality</label>
          <input type="text" id="nationality" class="form-control" value="Vietnamese" readonly>
        </div>
        <div class="form-group">
          <label>Academic Status</label>
          <select id="status" class="form-control" required>
            <option value="Enrolled">Enrolled</option>
            <option value="Graduated">Graduated</option>
            <option value="On Temporary Leave">On Temporary Leave</option>
            <option value="Dropped">Dropped</option>
          </select>
        </div>
        <div class="form-group">
          <label>Type of Visa</label>
          <input type="text" id="visaType" class="form-control" placeholder="e.g., 9(f) Student Visa" required>
        </div>
        <div class="form-group">
          <label>Passport Name</label>
          <input type="text" id="passportName" class="form-control" required>
        </div>
        <div class="form-group">
          <label>English Name</label>
          <input type="text" id="englishName" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Passport Number</label>
          <input type="text" id="passportNumber" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Birthday</label>
          <input type="date" id="birthday" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Civil Status</label>
          <select id="civilStatus" class="form-control" required>
            <option value="Single">Single</option>
            <option value="Married">Married</option>
            <option value="Divorced">Divorced</option>
            <option value="Widowed">Widowed</option>
          </select>
        </div>
        <div class="form-group">
          <label>Sex</label>
          <select id="sex" class="form-control" required>
            <option value="Female">Female</option>
            <option value="Male">Male</option>
          </select>
        </div>
        <div class="form-group full-width">
          <label>Address (Province/City in Vietnam)</label>
          <input type="text" id="address" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Course</label>
          <select id="course" class="form-control" required>
            <option value="Doctor of Business Administration (DBA)">Doctor of Business Administration (DBA)</option>
            <option value="Doctor of Philosophy in Educational Management (PhD-EdM)">Doctor of Philosophy in Educational Management (PhD-EdM)</option>
          </select>
        </div>
        <div class="form-group">
          <label>Admission Date</label>
          <input type="date" id="admissionDate" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Visa Registration Date</label>
          <input type="date" id="visaRegDate" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Visa Expiration Date</label>
          <input type="date" id="visaExpDate" class="form-control" required>
        </div>
        <div class="form-group full-width">
          <label>ACR I-CARD Number</label>
          <input type="text" id="acrNumber" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Last Arrival Date</label>
          <input type="date" id="arrivalDate" class="form-control" required>
        </div>
        <div class="form-group">
          <label>Last Departure Date</label>
          <input type="date" id="departureDate" class="form-control">
        </div>
        <div class="form-group full-width">
          <label>Schedule for Immigration Annual Reporting</label>
          <input type="date" id="annualReporting" class="form-control" required>
        </div>

        <div class="form-group full-width">
          <label>Upload Photo from Device / Gallery</label>
          <input type="file" id="photoFileInput" class="form-control" accept="image/*" onchange="previewFilePhoto(event)">
          <div class="photo-preview-box">
            <img id="photoPreview" class="photo-preview-img" alt="Preview">
            <span id="photoPreviewText" style="font-size: 12px; color: var(--text-gray);">No photo selected</span>
          </div>
        </div>
      </div>
      <button type="submit" class="btn" style="width: 100%; margin-top: 20px;">Save Record</button>
    </form>
  </div>
</div>

<script>
// ADMIN CONFIGURATION
const ADMIN_EMAIL = "banaganhazelquennie14@gmail.com";

// USER DATABASE INITIALIZATION
let registeredUsers = JSON.parse(localStorage.getItem('iaelo_users')) || [
  {
    name: "Hazel Quennie Banagan",
    email: "banaganhazelquennie14@gmail.com",
    pass: "0123456789",
    role: "Admin"
  }
];

// PENDING REQUESTS INITIALIZATION
let pendingRequests = JSON.parse(localStorage.getItem('iaelo_pending_requests')) || [];

// Active selection tracker & Filter state
let activeStudentId = null;
let quickAlertFilter = 'none'; // 'visa', 'threshold', or 'none'

let students = JSON.parse(localStorage.getItem('iaelo_vietnamese_students')) || [
  {
    id: 1,
    nationality: "Vietnamese",
    status: "Enrolled",
    visaType: "9(f) Student Visa",
    passportName: "NGUYEN VAN A",
    englishName: "Alex Nguyen",
    passportNumber: "C1234567",
    birthday: "1992-05-14",
    civilStatus: "Single",
    sex: "Male",
    address: "Hanoi City",
    course: "Doctor of Business Administration (DBA)",
    admissionDate: "2025-07-15",
    visaRegDate: "2025-08-10",
    visaExpDate: "2026-08-10",
    acrNumber: "ACR987654321",
    arrivalDate: "2026-01-15",
    departureDate: "",
    annualReporting: "2027-02-14",
    photoUrl: "https://images.unsplash.com/photo-1534528741775-53994a69daeb?auto=format&fit=crop&w=300&q=80"
  }
];

window.onload = function() {
  checkAuth();
  loadSavedLogo();
};

function toggleAuthMode(mode) {
  const loginForm = document.getElementById('loginForm');
  const regForm = document.getElementById('registerForm');
  const heading = document.getElementById('authHeading');
  const subheading = document.getElementById('authSubheading');
  const msgBox = document.getElementById('authMessage');

  msgBox.style.display = 'none';

  if (mode === 'register') {
    loginForm.style.display = 'none';
    regForm.style.display = 'block';
    heading.textContent = "Request Access";
    subheading.textContent = "Submit your credentials for administrator review";
  } else {
    regForm.style.display = 'none';
    loginForm.style.display = 'block';
    heading.textContent = "IAELO Portal Access";
    subheading.textContent = "Enter your credentials or request access";
  }
}

function checkAuth() {
  const currentSession = JSON.parse(sessionStorage.getItem('iaelo_active_user'));
  if (currentSession) {
    document.getElementById('loginOverlay').style.display = 'none';
    
    const isAdmin = currentSession.email.toLowerCase() === ADMIN_EMAIL.toLowerCase();
    
    document.getElementById('currentUserDisplay').innerHTML = `
      Role: ${currentSession.role || (isAdmin ? 'Admin' : 'Viewer')}
    `;

    const canEdit = isAdmin || currentSession.role === 'Editor';

    document.querySelectorAll('.edit-allowed').forEach(el => {
      el.style.display = canEdit ? 'inline-flex' : 'none';
    });

    renderPendingRequests();
    renderStudents();
  } else {
    document.getElementById('loginOverlay').style.display = 'flex';
  }
}

function handleLogin(e) {
  e.preventDefault();
  const email = document.getElementById('loginEmail').value.trim().toLowerCase();
  const pass = document.getElementById('loginPassword').value;
  const msgBox = document.getElementById('authMessage');

  const user = registeredUsers.find(u => u.email.toLowerCase() === email && u.pass === pass);

  if (user) {
    sessionStorage.setItem('iaelo_active_user', JSON.stringify(user));
    msgBox.style.display = 'none';
    checkAuth();
  } else {
    msgBox.className = "auth-msg error";
    msgBox.textContent = "Invalid Email/Password, or your account is still pending approval.";
    msgBox.style.display = "block";
  }
}

function handleRegister(e) {
  e.preventDefault();
  const name = document.getElementById('regName').value.trim();
  const email = document.getElementById('regEmail').value.trim().toLowerCase();
  const pass = document.getElementById('regPassword').value;
  const confirmPass = document.getElementById('regConfirmPassword').value;
  const msgBox = document.getElementById('authMessage');

  if (pass !== confirmPass) {
    msgBox.className = "auth-msg error";
    msgBox.textContent = "Passwords do not match!";
    msgBox.style.display = "block";
    return;
  }

  if (email === ADMIN_EMAIL.toLowerCase()) {
    const adminUser = { name, email, pass, role: "Admin" };
    if (!registeredUsers.some(u => u.email.toLowerCase() === email)) {
      registeredUsers.push(adminUser);
      localStorage.setItem('iaelo_users', JSON.stringify(registeredUsers));
    }
    sessionStorage.setItem('iaelo_active_user', JSON.stringify(adminUser));
    checkAuth();
    return;
  }

  const existingUser = registeredUsers.find(u => u.email.toLowerCase() === email);
  if (existingUser) {
    msgBox.className = "auth-msg error";
    msgBox.textContent = "An account with this email already exists!";
    msgBox.style.display = "block";
    return;
  }

  const newReq = { name, email, pass };
  pendingRequests.push(newReq);
  localStorage.setItem('iaelo_pending_requests', JSON.stringify(pendingRequests));

  msgBox.className = "auth-msg success";
  msgBox.textContent = "Access request submitted successfully! Waiting for admin approval.";
  msgBox.style.display = "block";
  document.getElementById('registerForm').reset();
}

function toggleNotifDropdown() {
  const dropdown = document.getElementById('notifDropdown');
  dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
}

function renderPendingRequests() {
  const container = document.getElementById('pendingRequestsContainer');
  const badge = document.getElementById('bellBadge');
  const currentSession = JSON.parse(sessionStorage.getItem('iaelo_active_user'));
  const isAdmin = currentSession && currentSession.email.toLowerCase() === ADMIN_EMAIL.toLowerCase();

  container.innerHTML = '';

  if (pendingRequests.length > 0) {
    badge.textContent = pendingRequests.length;
    badge.style.display = 'inline-block';
  } else {
    badge.style.display = 'none';
  }

  if (pendingRequests.length === 0) {
    container.innerHTML = `<p style="font-size: 12px; color: var(--text-gray);">No pending access requests.</p>`;
    return;
  }

  pendingRequests.forEach((req, index) => {
    const card = document.createElement('div');
    card.className = 'request-card';
    
    if (isAdmin) {
      card.innerHTML = `
        <div class="request-info">
          <p>${req.name} (${req.email})</p>
          <span>Appoint Access Level:</span>
        </div>
        <div class="request-actions">
          <button class="btn btn-secondary" style="padding: 5px 10px; font-size: 11px;" onclick="approveRequest(${index}, 'Viewer')">👁️ View Only</button>
          <button class="btn" style="padding: 5px 10px; font-size: 11px;" onclick="approveRequest(${index}, 'Editor')">✏️ Can Edit</button>
          <button class="btn btn-danger" style="padding: 5px 10px; font-size: 11px;" onclick="rejectRequest(${index})">Reject</button>
        </div>
      `;
    } else {
      card.innerHTML = `
        <div class="request-info">
          <p>${req.name}</p>
          <span>Request Status: <strong>Pending Approval</strong></span>
        </div>
      `;
    }
    container.appendChild(card);
  });
}

function approveRequest(index, assignedRole) {
  const req = pendingRequests[index];
  const approvedUser = {
    name: req.name,
    email: req.email,
    pass: req.pass,
    role: assignedRole
  };

  registeredUsers.push(approvedUser);
  localStorage.setItem('iaelo_users', JSON.stringify(registeredUsers));

  pendingRequests.splice(index, 1);
  localStorage.setItem('iaelo_pending_requests', JSON.stringify(pendingRequests));

  renderPendingRequests();
  alert(`Successfully approved ${req.name} with ${assignedRole === 'Viewer' ? 'View-Only' : 'Can Edit'} access!`);
}

function rejectRequest(index) {
  if (confirm("Are you sure you want to reject this request?")) {
    pendingRequests.splice(index, 1);
    localStorage.setItem('iaelo_pending_requests', JSON.stringify(pendingRequests));
    renderPendingRequests();
  }
}

function openLogoModal() {
  const currentLogo = document.getElementById('appLogo').src;
  document.getElementById('modalLogoPreview').src = currentLogo;
  document.getElementById('logoModal').style.display = 'flex';
}

function closeLogoModal() {
  document.getElementById('logoModal').style.display = 'none';
}

function handleLogoUpload(event) {
  const file = event.target.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = function(e) {
      const logoBase64 = e.target.result;
      document.getElementById('appLogo').src = logoBase64;
      document.getElementById('modalLogoPreview').src = logoBase64;
      localStorage.setItem('iaelo_app_logo', logoBase64);
      alert("Logo updated successfully!");
    };
    reader.readAsDataURL(file);
  }
}

function loadSavedLogo() {
  const savedLogo = localStorage.getItem('iaelo_app_logo');
  if (savedLogo) {
    document.getElementById('appLogo').src = savedLogo;
  }
}

function handleLogout() {
  sessionStorage.removeItem('iaelo_active_user');
  document.getElementById('loginEmail').value = '';
  document.getElementById('loginPassword').value = '';
  toggleAuthMode('login');
  checkAuth();
}

function saveData() {
  localStorage.setItem('iaelo_vietnamese_students', JSON.stringify(students));
}

function previewFilePhoto(event) {
  const file = event.target.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = function(e) {
      const base64Image = e.target.result;
      document.getElementById('currentPhotoBase64').value = base64Image;
      const imgPreview = document.getElementById('photoPreview');
      imgPreview.src = base64Image;
      imgPreview.style.display = 'block';
      document.getElementById('photoPreviewText').textContent = file.name;
    };
    reader.readAsDataURL(file);
  }
}

function getStatusBadgeClass(status) {
  switch(status) {
    case 'Enrolled': return 'alert-success';
    case 'Graduated': return 'alert-info';
    case 'On Temporary Leave': return 'alert-warning';
    case 'Dropped': return 'alert-danger';
    default: return 'alert-success';
  }
}

function setQuickFilter(alertType) {
  if (alertType === 'all') {
    quickAlertFilter = 'none';
  } else if (alertType === 'visa-alerts') {
    quickAlertFilter = 'visa';
  } else if (alertType === 'threshold-alerts') {
    quickAlertFilter = 'threshold';
  }
  renderStudents();
}

function renderStudents() {
  const container = document.getElementById('studentListContainer');
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  container.innerHTML = '';

  let visaAlerts = 0;
  let thresholdAlerts = 0;
  const now = new Date();

  // Calculate overall metrics
  students.forEach(student => {
    const arrivalDate = new Date(student.arrivalDate);
    const sixMonthsAfterArrival = new Date(arrivalDate);
    sixMonthsAfterArrival.setMonth(sixMonthsAfterArrival.getMonth() + 6);
    const diffTime = sixMonthsAfterArrival - now;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    if (diffDays <= 30) thresholdAlerts++;

    const visaExp = new Date(student.visaExpDate);
    const visaDiffTime = visaExp - now;
    const visaDiffDays = Math.ceil(visaDiffTime / (1000 * 60 * 60 * 24));
    if (visaDiffDays <= 60 && visaDiffDays >= 0) visaAlerts++;
  });

  document.getElementById('totalStudentsCount').textContent = students.length;
  document.getElementById('visaAlertsCount').textContent = visaAlerts;
  document.getElementById('thresholdAlertsCount').textContent = thresholdAlerts;

  // Filter List Data via Search or Stat Card clicks
  const filtered = students.filter(s => {
    const matchesSearch = s.passportName.toLowerCase().includes(searchTerm) ||
      s.englishName.toLowerCase().includes(searchTerm) ||
      s.course.toLowerCase().includes(searchTerm) ||
      (s.status && s.status.toLowerCase().includes(searchTerm)) ||
      s.passportNumber.toLowerCase().includes(searchTerm);

    let matchesQuickAlert = true;
    if (quickAlertFilter === 'visa') {
      const visaExp = new Date(s.visaExpDate);
      const visaDiffDays = Math.ceil((visaExp - now) / (1000 * 60 * 60 * 24));
      matchesQuickAlert = (visaDiffDays <= 60 && visaDiffDays >= 0);
    } else if (quickAlertFilter === 'threshold') {
      const arrivalDate = new Date(s.arrivalDate);
      const sixMonthsAfterArrival = new Date(arrivalDate);
      sixMonthsAfterArrival.setMonth(sixMonthsAfterArrival.getMonth() + 6);
      const diffDays = Math.ceil((sixMonthsAfterArrival - now) / (1000 * 60 * 60 * 24));
      matchesQuickAlert = (diffDays <= 30);
    }

    return matchesSearch && matchesQuickAlert;
  });

  if (filtered.length === 0) {
    container.innerHTML = `<div style="padding: 24px; text-align: center; color: var(--text-gray); font-weight: 600;">No Vietnamese student records match your selected criteria.</div>`;
    return;
  }

  filtered.forEach(student => {
    const item = document.createElement('div');
    const isSelected = student.id === activeStudentId;
    item.className = `student-list-item ${isSelected ? 'active-student' : ''}`;
    
    item.onclick = () => selectStudent(student.id);
    item.innerHTML = `
      <div class="student-list-left">
        <img src="${student.photoUrl || 'https://via.placeholder.com/50'}" class="student-thumb" alt="Photo" onerror="this.src='https://via.placeholder.com/50'">
        <div class="student-basic-info">
          <h4>${student.englishName} (${student.passportName})</h4>
          <p>${student.course} • Passport: ${student.passportNumber} • <strong style="color: var(--primary);">${student.status || 'Enrolled'}</strong></p>
        </div>
      </div>
      <div class="click-hint">${isSelected ? 'Active Selection ✓' : 'View Profile & Compliance ➔'}</div>
    `;
    container.appendChild(item);
  });
}

function selectStudent(id) {
  activeStudentId = id;
  renderStudents();
  openViewModal(id);
}

function openViewModal(id) {
  const student = students.find(s => s.id === id);
  if (!student) return;

  const now = new Date();
  const arrivalDate = new Date(student.arrivalDate);
  const sixMonthsAfterArrival = new Date(arrivalDate);
  sixMonthsAfterArrival.setMonth(sixMonthsAfterArrival.getMonth() + 6);
  const diffTime = sixMonthsAfterArrival - now;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  let thresholdStatus = `<span class="alert-badge alert-success">Within 6-Month Entry Window</span>`;
  if (diffDays <= 30 && diffDays >= 0) {
    thresholdStatus = `<span class="alert-badge alert-warning">Approaching 6-Month Mark (${diffDays} days left)</span>`;
  } else if (diffDays < 0) {
    thresholdStatus = `<span class="alert-badge alert-danger">Exceeded 6 Months from Last Entry!</span>`;
  }

  const statusClass = getStatusBadgeClass(student.status || 'Enrolled');

  document.getElementById('viewModalTitle').textContent = student.englishName + " - Profile & Compliance";

  const modalBody = document.getElementById('viewModalBody');
  modalBody.innerHTML = `
    <div class="profile-view-header">
      <img src="${student.photoUrl || 'https://via.placeholder.com/85'}" class="profile-view-img" alt="Photo" onerror="this.src='https://via.placeholder.com/85'">
      <div>
        <h4 style="color: var(--primary); font-size: 18px; margin-bottom: 2px; font-weight: 800;">${student.englishName}</h4>
        <p style="color: var(--text-gray); font-size: 13px; font-weight: 600;">${student.passportName}</p>
        <div style="margin-top: 6px; display: flex; gap: 6px;">
          <span class="alert-badge alert-success">${student.nationality}</span>
          <span class="alert-badge ${statusClass}">${student.status || 'Enrolled'}</span>
        </div>
      </div>
    </div>
    <div class="info-row"><span>Academic Status:</span> <span>${student.status || 'Enrolled'}</span></div>
    <div class="info-row"><span>Course:</span> <span>${student.course}</span></div>
    <div class="info-row"><span>Admission Date:</span> <span>${student.admissionDate || 'N/A'}</span></div>
    <div class="info-row"><span>Passport Number:</span> <span>${student.passportNumber}</span></div>
    <div class="info-row"><span>Birthday:</span> <span>${student.birthday}</span></div>
    <div class="info-row"><span>Civil Status:</span> <span>${student.civilStatus}</span></div>
    <div class="info-row"><span>Sex:</span> <span>${student.sex}</span></div>
    <div class="info-row"><span>Address in Vietnam:</span> <span>${student.address}</span></div>
    <div class="info-row"><span>Type of Visa:</span> <span>${student.visaType || 'N/A'}</span></div>
    <div class="info-row"><span>Visa Registration Date:</span> <span>${student.visaRegDate}</span></div>
    <div class="info-row"><span>Visa Expiration Date:</span> <span>${student.visaExpDate}</span></div>
    <div class="info-row"><span>ACR I-CARD Number:</span> <span>${student.acrNumber}</span></div>

    <div class="compliance-box">
      <h5>Compliance Tracker</h5>
      <div class="info-row" style="border:none; margin-bottom:4px;"><span>Last Arrival Date:</span> <span>${student.arrivalDate}</span></div>
      <div class="info-row" style="border:none; margin-bottom:4px;"><span>Last Departure Date:</span> <span>${student.departureDate || 'None recorded'}</span></div>
      <div class="info-row" style="border:none; margin-bottom:4px;"><span>Immigration Annual Reporting:</span> <span>${student.annualReporting}</span></div>
      <div style="text-align: right; margin-top: 8px;">${thresholdStatus}</div>
    </div>
  `;

  document.getElementById('editFromViewBtn').onclick = function() {
    closeViewModal();
    openEditModal(student.id);
  };

  document.getElementById('deleteFromViewBtn').onclick = function() {
    deleteStudent(student.id);
  };

  document.getElementById('viewModal').style.display = 'flex';
}

function closeViewModal() {
  document.getElementById('viewModal').style.display = 'none';
}

function openAddModal() {
  document.getElementById('modalTitle').textContent = "Add Vietnamese Student Record";
  document.getElementById('studentForm').reset();
  document.getElementById('studentId').value = '';
  document.getElementById('currentPhotoBase64').value = '';
  document.getElementById('nationality').value = "Vietnamese";
  document.getElementById('status').value = "Enrolled";
  document.getElementById('civilStatus').value = "Single";
  document.getElementById('photoPreview').style.display = 'none';
  document.getElementById('photoPreviewText').textContent = 'No photo selected';
  document.getElementById('studentModal').style.display = 'flex';
}

function openEditModal(id) {
  document.getElementById('modalTitle').textContent = "Edit Student Record";
  const student = students.find(s => s.id === id);
  if (!student) return;

  document.getElementById('studentId').value = student.id;
  document.getElementById('nationality').value = "Vietnamese";
  document.getElementById('status').value = student.status || 'Enrolled';
  document.getElementById('visaType').value = student.visaType || '';
  document.getElementById('passportName').value = student.passportName;
  document.getElementById('englishName').value = student.englishName;
  document.getElementById('passportNumber').value = student.passportNumber;
  document.getElementById('birthday').value = student.birthday;
  
  let statusVal = student.civilStatus;
  if (statusVal === 'Widow') statusVal = 'Widowed';
  document.getElementById('civilStatus').value = statusVal || 'Single';

  document.getElementById('sex').value = student.sex;
  document.getElementById('address').value = student.address;
  document.getElementById('course').value = student.course;
  document.getElementById('admissionDate').value = student.admissionDate || '';
  document.getElementById('visaRegDate').value = student.visaRegDate;
  document.getElementById('visaExpDate').value = student.visaExpDate;
  document.getElementById('acrNumber').value = student.acrNumber;
  document.getElementById('arrivalDate').value = student.arrivalDate;
  document.getElementById('departureDate').value = student.departureDate;
  document.getElementById('annualReporting').value = student.annualReporting;

  document.getElementById('currentPhotoBase64').value = student.photoUrl || '';
  if (student.photoUrl) {
    const imgPreview = document.getElementById('photoPreview');
    imgPreview.src = student.photoUrl;
    imgPreview.style.display = 'block';
    document.getElementById('photoPreviewText').textContent = 'Existing photo active';
  } else {
    document.getElementById('photoPreview').style.display = 'none';
    document.getElementById('photoPreviewText').textContent = 'No photo selected';
  }

  document.getElementById('studentModal').style.display = 'flex';
}

function closeModal() {
  document.getElementById('studentModal').style.display = 'none';
}

function saveStudent(e) {
  e.preventDefault();
  const id = document.getElementById('studentId').value;
  
  const studentData = {
    id: id ? parseInt(id) : Date.now(),
    nationality: document.getElementById('nationality').value,
    status: document.getElementById('status').value,
    visaType: document.getElementById('visaType').value,
    passportName: document.getElementById('passportName').value.toUpperCase(),
    englishName: document.getElementById('englishName').value,
    passportNumber: document.getElementById('passportNumber').value,
    birthday: document.getElementById('birthday').value,
    civilStatus: document.getElementById('civilStatus').value,
    sex: document.getElementById('sex').value,
    address: document.getElementById('address').value,
    course: document.getElementById('course').value,
    admissionDate: document.getElementById('admissionDate').value,
    visaRegDate: document.getElementById('visaRegDate').value,
    visaExpDate: document.getElementById('visaExpDate').value,
    acrNumber: document.getElementById('acrNumber').value,
    arrivalDate: document.getElementById('arrivalDate').value,
    departureDate: document.getElementById('departureDate').value,
    annualReporting: document.getElementById('annualReporting').value,
    photoUrl: document.getElementById('currentPhotoBase64').value || 'https://via.placeholder.com/100'
  };

  if (id) {
    const index = students.findIndex(s => s.id === parseInt(id));
    if (index !== -1) students[index] = studentData;
  } else {
    students.push(studentData);
    activeStudentId = studentData.id;
  }

  saveData();
  renderStudents();
  closeModal();
}

function deleteStudent(id) {
  if (confirm("Are you sure you want to delete this student record?")) {
    students = students.filter(s => s.id !== id);
    if (activeStudentId === id) activeStudentId = null;
    saveData();
    renderStudents();
    closeViewModal();
  }
}

/* EXCEL IMPORT / EXPORT LOGIC */
function exportToExcel() {
  const exportData = students.map(s => ({
    "Nationality": s.nationality,
    "Status": s.status || "Enrolled",
    "Type of Visa": s.visaType,
    "Passport Name": s.passportName,
    "English Name": s.englishName,
    "Passport Number": s.passportNumber,
    "Birthday": s.birthday,
    "Civil Status": s.civilStatus === 'Widow' ? 'Widowed' : s.civilStatus,
    "Sex": s.sex,
    "Address in Vietnam": s.address,
    "Course": s.course,
    "Admission Date": s.admissionDate,
    "Visa Registration Date": s.visaRegDate,
    "Visa Expiration Date": s.visaExpDate,
    "ACR I-CARD Number": s.acrNumber,
    "Last Arrival Date": s.arrivalDate,
    "Last Departure Date": s.departureDate,
    "Immigration Annual Reporting": s.annualReporting
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Vietnamese Students");
  XLSX.writeFile(workbook, "IAELO_Vietnamese_Students_Directory.xlsx");
}

function importFromExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

    if (jsonData.length === 0) {
      alert("No valid data found in Excel file.");
      return;
    }

    jsonData.forEach((row, idx) => {
      let importedCivil = row["Civil Status"] || "Single";
      if (importedCivil === 'Widow') importedCivil = 'Widowed';

      students.push({
        id: Date.now() + idx,
        nationality: row["Nationality"] || "Vietnamese",
        status: row["Status"] || "Enrolled",
        visaType: row["Type of Visa"] || "9(f) Student Visa",
        passportName: row["Passport Name"] || "N/A",
        englishName: row["English Name"] || "N/A",
        passportNumber: row["Passport Number"] || "N/A",
        birthday: row["Birthday"] || "",
        civilStatus: importedCivil,
        sex: row["Sex"] || "Female",
        address: row["Address in Vietnam"] || "",
        course: row["Course"] || "Doctor of Business Administration (DBA)",
        admissionDate: row["Admission Date"] || "",
        visaRegDate: row["Visa Registration Date"] || "",
        visaExpDate: row["Visa Expiration Date"] || "",
        acrNumber: row["ACR I-CARD Number"] || "",
        arrivalDate: row["Last Arrival Date"] || "",
        departureDate: row["Last Departure Date"] || row["Departure Date"] || "",
        annualReporting: row["Immigration Annual Reporting"] || "",
        photoUrl: ""
      });
    });

    saveData();
    renderStudents();
    alert(`Successfully imported ${jsonData.length} student records!`);
  };
  reader.readAsArrayBuffer(file);
}
</script>
</body>
</html>
