const els = {
  authGate: document.getElementById("auth-gate"),
  appShell: document.getElementById("app-shell"),
  accessPassword: document.getElementById("access-password"),
  accessSubmit: document.getElementById("access-submit"),
  accessError: document.getElementById("access-error"),
  homeView: document.getElementById("home-view"),
  dashboardView: document.getElementById("dashboard-view"),
  homeUpload: document.getElementById("home-upload"),
  fileInput: document.getElementById("file-input"),
  status: document.getElementById("status"),
  workbookMeta: document.getElementById("workbook-meta"),
  reportLink: document.getElementById("report-link"),
  kpiGrid: document.getElementById("kpi-grid"),
  riskList: document.getElementById("risk-list"),
  migrationSummary: document.getElementById("migration-summary"),
  migrationActions: document.getElementById("migration-actions"),
  lifecycleSummary: document.getElementById("lifecycle-summary"),
  barReady: document.getElementById("bar-ready"),
  dataInfraSummary: document.getElementById("data-infra-summary"),
  networkSummary: document.getElementById("network-summary"),
  criticalCounts: document.getElementById("critical-counts"),
  principlesNote: document.getElementById("principles-note"),
  clientSummary: document.getElementById("client-summary"),
  cyberSummary: document.getElementById("cyber-summary"),
  coreSummary: document.getElementById("core-summary"),
  brilliantBasicsSummary: document.getElementById("brilliant-basics-summary"),
  sheetSelect: document.getElementById("sheet-select"),
  sheetInfo: document.getElementById("sheet-info"),
  sheetTable: document.getElementById("sheet-table"),
  exportWeb: document.getElementById("export-web"),
  exportWord: document.getElementById("export-word"),
  exportPdf: document.getElementById("export-pdf"),
};

let workbook = null;
let latestReport = null;
const ACCESS_STATE_KEY = "audit-service-unlocked";
const ACCESS_HASH_HEX = "cff2083f7f003ef5b8db6b8c8b43abaf68188c6ed1ea02d5d6b477f789d06ab0";

els.fileInput.addEventListener("click", () => {
  els.fileInput.value = "";
});
els.fileInput.addEventListener("change", onUpload);
if (els.homeUpload) {
  els.homeUpload.addEventListener("click", () => {
    els.fileInput.value = "";
    els.fileInput.click();
  });
}
els.sheetSelect.addEventListener("change", renderSheetTable);
els.exportWeb.addEventListener("click", exportWeb);
els.exportWord.addEventListener("click", exportWord);
els.exportPdf.addEventListener("click", exportPdf);
if (els.accessSubmit) els.accessSubmit.addEventListener("click", unlockAccess);
if (els.accessPassword) {
  els.accessPassword.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      unlockAccess();
    }
  });
}
initAccessGate();

function setAccessState(unlocked) {
  if (els.authGate) els.authGate.classList.toggle("hidden", unlocked);
  if (els.appShell) els.appShell.classList.toggle("hidden", !unlocked);
}

function initAccessGate() {
  const unlocked = sessionStorage.getItem(ACCESS_STATE_KEY) === "1";
  setAccessState(unlocked);
}

async function unlockAccess() {
  if (!els.accessPassword) return;
  const entered = els.accessPassword.value || "";
  const hash = await hashTextSha256(entered);
  if (hash === ACCESS_HASH_HEX) {
    sessionStorage.setItem(ACCESS_STATE_KEY, "1");
    if (els.accessError) els.accessError.classList.add("hidden");
    els.accessPassword.value = "";
    setAccessState(true);
    return;
  }
  if (els.accessError) els.accessError.classList.remove("hidden");
}

async function hashTextSha256(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const arr = Array.from(new Uint8Array(digest));
  return arr.map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function onUpload(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const arrayBuffer = await file.arrayBuffer();
  workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  showDashboardView();
  els.status.textContent = "Workbook loaded.";
  const modifiedAt = file.lastModified ? new Date(file.lastModified).toLocaleString() : "Unknown";
  const loadedAt = new Date().toLocaleString();
  const sizeKb = Math.max(1, Math.round(file.size / 1024));
  els.workbookMeta.textContent = `${file.name} | ${sizeKb} KB | Modified: ${modifiedAt} | Loaded: ${loadedAt} | Sheets: ${workbook.SheetNames.length}`;
  if (els.reportLink) els.reportLink.textContent = "";
  els.exportWeb.disabled = false;
  els.exportWord.disabled = false;
  els.exportPdf.disabled = false;

  renderDashboard();
  populateSheetSelect();
  renderSheetTable();
  event.target.value = "";
}

function showDashboardView() {
  if (els.homeView) els.homeView.classList.add("hidden");
  if (els.dashboardView) els.dashboardView.classList.remove("hidden");
}

function ensureBrilliantBasicsContainer() {
  if (els.brilliantBasicsSummary) return;
  const dashboard = els.dashboardView;
  if (!dashboard) return;

  const sheetExplorer = els.sheetSelect ? els.sheetSelect.closest(".card") : null;
  const section = document.createElement("section");
  section.className = "card";
  section.innerHTML = `
    <h2>Unleashed Brilliant Basics Alignment</h2>
    <div id="brilliant-basics-summary"></div>
  `;

  if (sheetExplorer && sheetExplorer.parentNode === dashboard) dashboard.insertBefore(section, sheetExplorer);
  else dashboard.appendChild(section);
  els.brilliantBasicsSummary = section.querySelector("#brilliant-basics-summary");
}

function renderVmOsStatus(vmOs) {
  if (!vmOs) return "";
  return `
    <details class="details-list lifecycle-dropdown" open>
      <summary><strong>Virtual machine OS status</strong></summary>
      <div class="lifecycle-columns">
        <div class="lifecycle-col">
          <div class="lifecycle-col-head">
            <strong>Flagged (bad status)</strong>
            <span>${vmOs.flagged.length}</span>
          </div>
          ${
            vmOs.flagged.length
              ? `<ul class="list lifecycle-list">${vmOs.flagged.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
              : `<p class="muted">No flagged virtual machine OS items.</p>`
          }
        </div>
        <div class="lifecycle-col">
          <div class="lifecycle-col-head">
            <strong>Good status</strong>
            <span>${vmOs.good.length}</span>
          </div>
          ${
            vmOs.good.length
              ? `<ul class="list lifecycle-list">${vmOs.good.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
              : `<p class="muted">No virtual machines in good status.</p>`
          }
        </div>
        <div class="lifecycle-col">
          <div class="lifecycle-col-head">
            <strong>Rule</strong>
            <span>Info</span>
          </div>
          <p class="muted">VM host lifecycle is inherited from physical servers. VM risk here is based on guest OS support status (for example, Windows Server 2012 is urgent).</p>
        </div>
      </div>
    </details>
  `;
}

function renderBoardColumn(title, toneClass, items, emptyText) {
  const filtered = (items || []).filter((item) => String(item || "").trim() !== "");
  return `
    <div class="migration-stage-card ${toneClass}">
      <div class="migration-stage-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${filtered.length}</span>
      </div>
      ${
        filtered.length
          ? `<ul class="list migration-stage-list">${filtered.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : `<p class="muted">${escapeHtml(emptyText)}</p>`
      }
    </div>
  `;
}

function lifecycleFlagSummary(lifecycle) {
  if (!lifecycle) return [];
  return [
    `End of warranty flagged: ${(lifecycle.flagged.warranty || []).length}`,
    `End of support flagged: ${(lifecycle.flagged.support || []).length}`,
    `End of service life flagged: ${(lifecycle.flagged.serviceLife || []).length}`,
  ];
}

function renderDashboard() {
  ensureBrilliantBasicsContainer();
  const dashboardRows = getRows("Dashboard");
  const infraRows = getRows("Data Infrastructure");
  const networkRows = getRows("Network & WiFi");
  const clientRows = getRows("Client Compute");
  const cyberRows = getRows("Cyber Security");
  const migrationRows = getRows("Migration Readiness");
  const coreRows = getRows("Core Application");
  const softwareRows = getRows("Software");
  const questionsRows = getRows("Questions to Confirm");

  const dashboard = parseDashboard(dashboardRows);
  const infra = parseDataInfrastructure(infraRows);
  const network = parseNetwork(networkRows);
  const client = parseClient(clientRows);
  const cyber = parseCyber(cyberRows);
  const migration = parseMigration(migrationRows);
  const core = parseCore(coreRows);
  const software = parseSoftware(softwareRows);
  const questions = parseQuestions(questionsRows);
  const cyberPctFallback = cyber.totalControls > 0
    ? `${((cyber.completeCount / cyber.totalControls) * 100).toFixed(1)}%`
    : "-";
  const cyberPctValue = dashboard.cyberPct !== "-" ? dashboard.cyberPct : cyberPctFallback;

  latestReport = { dashboard, infra, network, client, cyber, migration, core, software, questions };

  renderKpis(dashboard, infra, network, client, cyber, migration, questions);
  renderRisks(infra, cyber, migration, client, core);
  renderMigrationActions(migration);
  renderLifecycleSummary(infra, network, client);
  renderCriticalCounts(infra, network, client);

  els.principlesNote.textContent =
    "Principles in use: 5-year hardware review cycle; support-end is a critical replacement trigger; cloud-first transition (SaaS, then Azure, local only where dependency requires).";

  const infraFlagList = lifecycleFlagSummary(infra.lifecycle);
  els.dataInfraSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric"><div class="label">Total server records</div><div class="value">${infra.totalServers}</div></div>
      <div class="metric"><div class="label">Physical / Virtual</div><div class="value">${infra.physical} / ${infra.virtual}</div></div>
      <div class="metric"><div class="label">Immediate server candidates</div><div class="value ${infra.serverCritical > 0 ? "danger" : "ok"}">${infra.serverCritical}</div></div>
      <div class="metric"><div class="label">Windows Server 2012</div><div class="value ${infra.ws2012 > 0 ? "danger" : "ok"}">${infra.ws2012}</div></div>
      <div class="metric"><div class="label">SAN switches / storage arrays</div><div class="value">${infra.sanSwitchCount} / ${infra.sanStorageCount}</div></div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${renderBoardColumn("Server lifecycle flags", "stage-remediation", infraFlagList, "No server lifecycle flags.")}
      ${renderBoardColumn("Server OS posture", "stage-progress", [
        `Windows Server 2012 workloads: ${infra.ws2012} (urgent)`,
        `Windows Server 2016 workloads: ${infra.ws2016} (nearer lifecycle edge)`,
        `Physical monitor/planning items: ${infra.physicalMonitor}`,
      ], "No server OS lifecycle concerns.")}
      ${renderBoardColumn("Planning notes", "stage-other", [
        "Review hardware every 5 years.",
        "Support-end is critical replacement trigger.",
        "Physical presence alone is not replacement trigger.",
      ], "No planning notes.")}
    </div>
    ${renderVmOsStatus(infra.vmOs)}
    ${renderLifecycleDetails(infra.lifecycle)}
  `;

  const networkFlagList = lifecycleFlagSummary(network.lifecycle);
  const additionalNetworkItems = Object.entries(network.additionalSections || {})
    .map(([name, count]) => `${name}: ${count}`);
  els.networkSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric"><div class="label">Core switches</div><div class="value">${network.core}</div></div>
      <div class="metric"><div class="label">Edge switches</div><div class="value">${network.edge}</div></div>
      <div class="metric"><div class="label">Firewalls</div><div class="value">${network.firewalls}</div></div>
      <div class="metric"><div class="label">Wi-Fi APs</div><div class="value">${network.aps}</div></div>
      <div class="metric"><div class="label">Wi-Fi 5 / Wi-Fi 6</div><div class="value">${network.wifi5} / ${network.wifi6}</div></div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${renderBoardColumn("Network lifecycle flags", "stage-progress", networkFlagList, "No network lifecycle flags.")}
      ${renderBoardColumn("Warranty hotspots", "stage-remediation", [
        `APs out of warranty: ${network.apsOutOfWarranty}`,
        `Switches out of warranty: ${network.switchesOutOfWarranty}`,
        `Firewalls out of warranty: ${network.firewallsOutOfWarranty}`,
      ], "No warranty hotspots.")}
      ${renderBoardColumn("Target architecture & discovered components", "stage-other", [
        "Full Wi-Fi 7 programme target.",
        "Core switching to 10Gb first.",
        "Edge switching to 2.5Gb in phased rollout.",
        ...additionalNetworkItems,
      ], "No target architecture notes.")}
    </div>
    ${renderLifecycleDetails(network.lifecycle)}
  `;

  const clientFlagList = lifecycleFlagSummary(client.lifecycle);
  els.clientSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric"><div class="label">Device rows</div><div class="value">${client.deviceRows}</div></div>
      <div class="metric"><div class="label">Total quantity</div><div class="value">${client.totalQuantity}</div></div>
      <div class="metric"><div class="label">Windows / Chromebook / Tablet / Mac / Other</div><div class="value">${client.windowsDevices} / ${client.chromebooks} / ${client.tabletDevices} / ${client.macDevices} / ${client.otherDevices}</div></div>
      <div class="metric"><div class="label">Old OS total (client)</div><div class="value ${client.oldOsTotal > 0 ? "danger" : "ok"}">${client.oldOsTotal}</div></div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${renderBoardColumn("Client lifecycle flags", "stage-progress", clientFlagList, "No client lifecycle flags.")}
      ${renderBoardColumn("OS risk hotspots", "stage-remediation", [
        `Windows 10 quantity: ${client.windows10}`,
        `Legacy Windows (7/8/8.1): ${client.legacyWindows}`,
        `ChromeOS devices marked expired: ${client.chromeExpired}`,
      ], "No OS risk hotspots.")}
      ${renderBoardColumn("Data quality notes", "stage-other", [
        `Tablet OS not specified: ${client.tabletOsUnknown}`,
        "Unsupported OS cohorts are immediate replacement candidates.",
      ], "No data quality notes.")}
    </div>
    ${renderLifecycleDetails(client.lifecycle)}
  `;

  const cyberCritical = [];
  if (isCyberGap(cyber.m365BackupStatus)) cyberCritical.push(`Microsoft 365 backup: ${cyber.m365BackupStatus}`);
  if (isCyberGap(cyber.emailSecurityStatus)) cyberCritical.push(`Email security: ${cyber.emailSecurityStatus}`);
  if (isCyberGap(cyber.cloudBackupStatus)) cyberCritical.push(`Cloud backup: ${cyber.cloudBackupStatus}`);
  const cyberDynamicGaps = (cyber.controls || [])
    .filter((c) => c.stage !== "complete")
    .slice(0, 80)
    .map((c) => `${c.area}: ${c.status}`);
  const cyberWorkstream = [
    `Not complete controls: ${cyber.incompleteCount}`,
    `N/A controls (not implemented): ${cyber.naCount}`,
    `Complete controls: ${cyber.completeCount}`,
  ];
  els.cyberSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric"><div class="label">Cyber completion</div><div class="value ${cyber.naCount > 0 ? "warn" : "ok"}">${cyberPctValue}</div></div>
      <div class="metric"><div class="label">Total controls</div><div class="value">${cyber.totalControls}</div></div>
      <div class="metric"><div class="label">Not complete controls</div><div class="value ${cyber.incompleteCount > 0 ? "danger" : "ok"}">${cyber.incompleteCount}</div></div>
      <div class="metric"><div class="label">N/A controls</div><div class="value ${cyber.naCount > 0 ? "danger" : "ok"}">${cyber.naCount}</div></div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${renderBoardColumn("Critical controls", "stage-remediation", cyberCritical, "No critical cyber controls flagged.")}
      ${renderBoardColumn("Control status", "stage-progress", cyberWorkstream, "No control status items.")}
      ${renderBoardColumn("Controls needing action (dynamic)", "stage-progress", cyberDynamicGaps, "No additional control gaps detected.")}
      ${renderBoardColumn("Programme notes", "stage-other", [
        "Why this %: reduced by not implemented/N/A controls.",
        "Interim Microsoft 365 backup and email security should be in place pre-migration.",
      ], "No programme notes.")}
    </div>
  `;

  els.coreSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric"><div class="label">Current licence</div><div class="value">${escapeHtml(core.currentLicence)}</div></div>
      <div class="metric"><div class="label">Feature rows analysed</div><div class="value">${core.featureRows}</div></div>
      <div class="metric"><div class="label">Configured but usage unknown</div><div class="value ${core.configuredUnknown > 0 ? "warn" : "ok"}">${core.configuredUnknown}</div></div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${renderBoardColumn("Optimisation priorities", "stage-progress", [
        `Configured but usage unknown: ${core.configuredUnknown}`,
        "Validate configured features against real adoption.",
      ], "No optimisation priorities.")}
      ${renderBoardColumn("Licensing direction", "stage-other", [
        "Use A3 baseline for current-state comparison.",
        "Map A5 + Acronis alignment to simplify vendor stack.",
      ], "No licensing direction notes.")}
      ${renderBoardColumn("Governance notes", "stage-other", [
        "Report excludes pricing by design.",
        "Focus on risk, simplification, and service outcomes.",
      ], "No governance notes.")}
    </div>
  `;

  const dashboardBasicsRows = buildBrilliantBasicsRows(cyber, core, software);
  const dashboardBasicsRowsHtml = dashboardBasicsRows
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.capability)}</td>
        <td><span class="basics-status ${escapeHtml(basicsStatusClassName(row.status))}">${escapeHtml(row.status)}</span></td>
        <td>${escapeHtml(row.evidence)}</td>
      </tr>
    `)
    .join("");

  if (els.brilliantBasicsSummary) {
    els.brilliantBasicsSummary.innerHTML = `
      <div class="table-wrap basics-table-wrap">
        <table class="basics-table">
          <thead><tr><th>Control</th><th>Status</th><th>Evidence note</th></tr></thead>
          <tbody>${dashboardBasicsRowsHtml}</tbody>
        </table>
      </div>
    `;
  }
}

function parseDashboard(rows) {
  const out = { servers: "-", edge: "-", wifi: "-", devices: "-", cyberPct: "-", msReadyPct: "-", msRemediation: "-" };
  for (const row of rows) {
    const k0 = normKey(row[0]);
    const k3 = normKey(row[3]);
    if (k3 === "servers") out.servers = row[4] ?? "-";
    if (k3 === "edge switches") out.edge = row[4] ?? "-";
    if (k3 === "wi-fi devices" || k3 === "wifi devices") out.wifi = row[4] ?? "-";
    if (k3 === "client devices") out.devices = row[4] ?? "-";
    if (
      (k0.includes("cyber") && (k0.includes("complete") || k0.includes("%") || k0.includes("percent"))) ||
      (k3.includes("cyber") && (k3.includes("complete") || k3.includes("%") || k3.includes("percent")))
    ) {
      const val = row[1] ?? row[4];
      const parsed = pct(val);
      if (parsed !== "-") out.cyberPct = parsed;
    }
    if (
      k0.includes("microsoft migration checks ready") ||
      k0.includes("microsoft readiness") ||
      k0.includes("migration readiness") ||
      k0.includes("ms readiness")
    ) out.msReadyPct = pct(row[1]);
    if (
      k0.includes("microsoft migration remediation items") ||
      k0.includes("remediation items") ||
      k0.includes("migration remediation")
    ) out.msRemediation = row[1] ?? "-";
  }
  return out;
}

function parseDataInfrastructure(rows) {
  let inServers = false;
  let section = "";
  let total = 0, physical = 0, virtual = 0, ws2012 = 0;
  let ws2016 = 0;
  let outOfWarranty = 0;
  let outOfServiceLife = 0;
  let outOfSupport = 0;
  let serverCritical = 0;
  let physicalMonitor = 0;
  let sanSwitchCount = 0;
  let sanStorageCount = 0;
  const lifecycle = createLifecycleStore();
  const vmOs = { flagged: [], good: [] };
  for (const row of rows) {
    if (row[0] === "Servers - Physical and Virtual") section = "servers";
    if (row[0] === "SAN Switching") section = "sanSwitch";
    if (row[0] === "SAN Storage") section = "sanStorage";

    if (row[0] === "Asset ID" && row[1] === "Server Name") { inServers = true; continue; }
    if (section === "servers" && inServers) {
      if (!row[1] || row[1] === "Server Name") continue;
      total += 1;
      const type = `${row[2] || ""}`.toLowerCase();
      const typeLabel = `${row[2] || "Unknown"}`;
      const isPhysical = type === "physical";
      const os = `${row[6] || ""}`.toLowerCase();
      const osRaw = `${row[6] || "OS unknown"}`;
      const warrantyStatus = `${row[9] || ""}`.toLowerCase();
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const warrantyExpiry = parseExcelDate(row[10]);
      const endOfSupportRaw = row[11];
      const endOfSupport = parseExcelDate(row[11]);
      const endOfServiceLifeRaw = row[12];
      const endOfServiceLife = parseExcelDate(row[12]);
      const itemLabel = `${row[1] || "Unnamed server"} (${typeLabel}, ${row[6] || "OS unknown"})`;
      let isCritical = false;
      if (type === "physical") physical += 1;
      if (type === "virtual") virtual += 1;
      if (os.includes("2012")) { ws2012 += 1; isCritical = true; }
      if (os.includes("2016")) ws2016 += 1;
      if (isPhysical && isPastDate(endOfSupport)) { isCritical = true; outOfSupport += 1; }
      if (isPhysical && isPastDate(endOfServiceLife)) outOfServiceLife += 1;
      if (isCritical) serverCritical += 1;
      if (isPhysical && isOutOfWarranty(warrantyStatus, warrantyExpiry)) outOfWarranty += 1;
      if (type === "physical" && !isCritical) physicalMonitor += 1;

      if (type === "virtual") {
        const vmLabel = `${row[1] || "Unnamed VM"} (${osRaw})`;
        if (isUnsupportedServerOs(osRaw)) {
          vmOs.flagged.push(`${vmLabel} - End of support / urgent upgrade`);
        } else {
          vmOs.good.push(`${vmLabel} - In supported OS cohort`);
        }
      }

      // VM lifecycle dates are not applicable here: lifecycle is inherited from the physical host.
      // For VMs, only OS state is evaluated (e.g. Server 2012).
      if (isPhysical) {
        addLifecycleEntry(
          lifecycle,
          "warranty",
          lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, warrantyExpiry),
          itemLabel
        );
        addLifecycleEntry(
          lifecycle,
          "support",
          lifecycleStatus("support", "", endOfSupportRaw, endOfSupport),
          itemLabel
        );
        addLifecycleEntry(
          lifecycle,
          "serviceLife",
          lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, endOfServiceLife),
          itemLabel
        );
      }
    }

    if (section === "sanSwitch") {
      if (row[0] === "Asset ID" && row[1] === "Switch Name") continue;
      if (row[1] && row[1] !== "Switch Name") sanSwitchCount += 1;
    }

    if (section === "sanStorage") {
      if (row[0] === "Asset ID" && row[1] === "Array Name") continue;
      if (row[1] && row[1] !== "Array Name") sanStorageCount += 1;
    }
  }
  return {
    totalServers: total,
    physical,
    virtual,
    ws2012,
    ws2016,
    outOfWarranty,
    outOfServiceLife,
    outOfSupport,
    serverCritical,
    physicalMonitor,
    sanSwitchCount,
    sanStorageCount,
    lifecycle,
    vmOs,
  };
}

function parseNetwork(rows) {
  let section = "";
  const knownSections = new Set(["Core Switches", "Edge Switches", "Wi-Fi", "Firewalls"]);
  let core = 0, edge = 0, aps = 0, wifi5 = 0, wifi6 = 0;
  let switchesOutOfWarranty = 0, apsOutOfWarranty = 0, firewallsOutOfWarranty = 0;
  let switchesOutOfServiceLife = 0, apsOutOfServiceLife = 0, firewallsOutOfServiceLife = 0;
  let switchesOutOfSupport = 0, apsOutOfSupport = 0, firewallsOutOfSupport = 0;
  let firewalls = 0;
  const lifecycle = createLifecycleStore();
  const additionalSections = {};
  for (const row of rows) {
    if (isSectionHeaderRow(row)) {
      section = String(row[0] || "").trim();
      continue;
    }
    if (row[0] === "Asset ID") continue;
    if (section === "Core Switches" && row[1]) {
      core += 1;
      const warrantyStatus = `${row[9] || ""}`.toLowerCase();
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const warrantyExpiry = parseExcelDate(row[10]);
      const endOfSupportRaw = row[12];
      const endOfSupport = parseExcelDate(row[12]);
      const endOfServiceLifeRaw = row[13];
      const endOfServiceLife = parseExcelDate(row[13]);
      const itemLabel = `${row[1]} (Core switch)`;
      if (isOutOfWarranty(warrantyStatus, warrantyExpiry)) switchesOutOfWarranty += 1;
      if (isPastDate(endOfSupport)) switchesOutOfSupport += 1;
      if (isPastDate(endOfServiceLife)) switchesOutOfServiceLife += 1;
      addLifecycleEntry(lifecycle, "warranty", lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, warrantyExpiry), itemLabel);
      addLifecycleEntry(lifecycle, "support", lifecycleStatus("support", "", endOfSupportRaw, endOfSupport), itemLabel);
      addLifecycleEntry(lifecycle, "serviceLife", lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, endOfServiceLife), itemLabel);
    }
    if (section === "Edge Switches" && row[1]) {
      edge += 1;
      const warrantyStatus = `${row[9] || ""}`.toLowerCase();
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const warrantyExpiry = parseExcelDate(row[10]);
      const endOfSupportRaw = row[12];
      const endOfSupport = parseExcelDate(row[12]);
      const endOfServiceLifeRaw = row[13];
      const endOfServiceLife = parseExcelDate(row[13]);
      const itemLabel = `${row[1]} (Edge switch)`;
      if (isOutOfWarranty(warrantyStatus, warrantyExpiry)) switchesOutOfWarranty += 1;
      if (isPastDate(endOfSupport)) switchesOutOfSupport += 1;
      if (isPastDate(endOfServiceLife)) switchesOutOfServiceLife += 1;
      addLifecycleEntry(lifecycle, "warranty", lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, warrantyExpiry), itemLabel);
      addLifecycleEntry(lifecycle, "support", lifecycleStatus("support", "", endOfSupportRaw, endOfSupport), itemLabel);
      addLifecycleEntry(lifecycle, "serviceLife", lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, endOfServiceLife), itemLabel);
    }
    if (section === "Firewalls" && row[1]) {
      firewalls += 1;
      const warrantyStatus = `${row[9] || ""}`.toLowerCase();
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const warrantyExpiry = parseExcelDate(row[10]);
      const endOfSupportRaw = row[13];
      const endOfSupport = parseExcelDate(row[13]);
      const endOfServiceLifeRaw = row[14];
      const endOfServiceLife = parseExcelDate(row[14]);
      const itemLabel = `${row[1]} (Firewall)`;
      if (isOutOfWarranty(warrantyStatus, warrantyExpiry)) firewallsOutOfWarranty += 1;
      if (isPastDate(endOfSupport)) firewallsOutOfSupport += 1;
      if (isPastDate(endOfServiceLife)) firewallsOutOfServiceLife += 1;
      addLifecycleEntry(lifecycle, "warranty", lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, warrantyExpiry), itemLabel);
      addLifecycleEntry(lifecycle, "support", lifecycleStatus("support", "", endOfSupportRaw, endOfSupport), itemLabel);
      addLifecycleEntry(lifecycle, "serviceLife", lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, endOfServiceLife), itemLabel);
    }
    if (section === "Wi-Fi" && row[1]) {
      aps += 1;
      const std = `${row[13] || ""}`.trim();
      if (std === "5") wifi5 += 1;
      if (std === "6") wifi6 += 1;
      const warrantyStatus = `${row[9] || ""}`.toLowerCase();
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const warrantyExpiry = parseExcelDate(row[10]);
      const endOfSupportRaw = row[14];
      const endOfSupport = parseExcelDate(row[14]);
      const endOfServiceLifeRaw = row[15];
      const endOfServiceLife = parseExcelDate(row[15]);
      const itemLabel = `${row[1]} (Wi-Fi AP)`;
      if (isOutOfWarranty(warrantyStatus, warrantyExpiry)) apsOutOfWarranty += 1;
      if (isPastDate(endOfSupport)) apsOutOfSupport += 1;
      if (isPastDate(endOfServiceLife)) apsOutOfServiceLife += 1;
      addLifecycleEntry(lifecycle, "warranty", lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, warrantyExpiry), itemLabel);
      addLifecycleEntry(lifecycle, "support", lifecycleStatus("support", "", endOfSupportRaw, endOfSupport), itemLabel);
      addLifecycleEntry(lifecycle, "serviceLife", lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, endOfServiceLife), itemLabel);
    }
    if (section && !knownSections.has(section) && row[1]) {
      additionalSections[section] = (additionalSections[section] || 0) + 1;
      const itemLabel = `${row[1]} (${section})`;
      const warrantyStatusRaw = `${row[9] || ""}`;
      const warrantyExpiryRaw = row[10];
      const endOfSupportRaw = row[12];
      const endOfServiceLifeRaw = row[13];
      addLifecycleEntry(lifecycle, "warranty", lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, parseExcelDate(warrantyExpiryRaw)), itemLabel);
      addLifecycleEntry(lifecycle, "support", lifecycleStatus("support", "", endOfSupportRaw, parseExcelDate(endOfSupportRaw)), itemLabel);
      addLifecycleEntry(lifecycle, "serviceLife", lifecycleStatus("serviceLife", "", endOfServiceLifeRaw, parseExcelDate(endOfServiceLifeRaw)), itemLabel);
    }
  }
  return {
    core,
    edge,
    aps,
    wifi5,
    wifi6,
    firewalls,
    switchesOutOfWarranty,
    apsOutOfWarranty,
    firewallsOutOfWarranty,
    switchesOutOfServiceLife,
    apsOutOfServiceLife,
    firewallsOutOfServiceLife,
    switchesOutOfSupport,
    apsOutOfSupport,
    firewallsOutOfSupport,
    lifecycle,
    additionalSections,
  };
}

function parseClient(rows) {
  let headerIdx = -1;
  let categoryIdx = 0, modelIdx = 1, osIdx = 3, qtyIdx = 4;
  let warrantyStatusIdx = -1, warrantyExpiryIdx = -1, endSupportIdx = -1, endServiceLifeIdx = -1, aueIdx = 9;

  for (let i = 0; i < rows.length; i++) {
    if (normKey(rows[i]?.[0]) === "device category") {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx >= 0) {
    const header = rows[headerIdx] || [];
    for (let i = 0; i < header.length; i++) {
      const key = normKey(header[i]);
      if (key === "device category") categoryIdx = i;
      if (key.includes("device name") || key.includes("model")) modelIdx = i;
      if (key === "operating system" || key === "os") osIdx = i;
      if (key === "quantity" || key.includes("qty")) qtyIdx = i;
      if (key.includes("warranty") && key.includes("status")) warrantyStatusIdx = i;
      if (key.includes("warranty") && (key.includes("expiry") || key.includes("end"))) warrantyExpiryIdx = i;
      if (key.includes("end of support") || key === "support end") endSupportIdx = i;
      if (key.includes("end of service") || key.includes("end of life")) endServiceLifeIdx = i;
      if (key.includes("aue") || key.includes("auto update expiry")) aueIdx = i;
    }
  }

  let deviceRows = 0, chromebooks = 0, windows10 = 0, legacyWindows = 0;
  let windowsDevices = 0, tabletDevices = 0, macDevices = 0, otherDevices = 0;
  let chromeExpired = 0, tabletOsUnknown = 0;
  let totalQuantity = 0;
  const lifecycle = createLifecycleStore();
  const start = headerIdx >= 0 ? headerIdx + 1 : 0;

  for (let i = start; i < rows.length; i++) {
    const row = rows[i] || [];
    if (!row[categoryIdx]) continue;
    deviceRows += 1;
    const categoryRaw = `${row[categoryIdx] || ""}`;
    const modelRaw = `${row[modelIdx] || ""}`;
    const osRaw = `${row[osIdx] || ""}`;
    const category = categoryRaw.toLowerCase();
    const model = modelRaw.toLowerCase();
    const os = osRaw.toLowerCase();
    const qty = Number(row[qtyIdx]) || 0;
    totalQuantity += qty;
    const labelQty = qty > 0 ? `x${qty}` : "x1";
    const itemLabel = `${modelRaw || categoryRaw || "Unnamed device"} ${labelQty}`.trim();

    if (os.includes("windows") || category.includes("windows")) windowsDevices += qty;
    else if (category.includes("chromebook") || os.includes("chromeos")) chromebooks += qty;
    else if (category.includes("tablet") || category.includes("ipad") || os.includes("ios") || os.includes("ipados") || os.includes("android")) tabletDevices += qty;
    else if (os.includes("mac") || os.includes("os x") || os.includes("macos") || category.includes("macbook") || model.includes("macbook") || model.includes("imac")) macDevices += qty;
    else otherDevices += qty;

    if (os.includes("windows 10")) windows10 += qty;
    if (os.includes("windows 7") || os.includes("windows 8")) legacyWindows += qty;
    if ((category.includes("chromebook") || os.includes("chromeos")) && String(row[aueIdx] || "").toLowerCase().includes("expired")) chromeExpired += qty;
    if ((category.includes("tablet") || category.includes("ipad")) && !os) tabletOsUnknown += qty;

    const warrantyStatusRaw = warrantyStatusIdx >= 0 ? row[warrantyStatusIdx] : "";
    const warrantyExpiryRaw = warrantyExpiryIdx >= 0 ? row[warrantyExpiryIdx] : "";
    const endSupportRaw = endSupportIdx >= 0 ? row[endSupportIdx] : "";
    const endServiceLifeRaw = endServiceLifeIdx >= 0 ? row[endServiceLifeIdx] : "";

    addLifecycleEntry(
      lifecycle,
      "warranty",
      lifecycleStatus("warranty", warrantyStatusRaw, warrantyExpiryRaw, parseExcelDate(warrantyExpiryRaw)),
      itemLabel
    );
    addLifecycleEntry(
      lifecycle,
      "support",
      lifecycleStatus("support", "", endSupportRaw, parseExcelDate(endSupportRaw)),
      itemLabel
    );
    addLifecycleEntry(
      lifecycle,
      "serviceLife",
      lifecycleStatus("serviceLife", "", endServiceLifeRaw, parseExcelDate(endServiceLifeRaw)),
      itemLabel
    );
  }
  return {
    deviceRows,
    totalQuantity,
    windowsDevices,
    chromebooks,
    tabletDevices,
    macDevices,
    otherDevices,
    windows10,
    legacyWindows,
    oldOsTotal: windows10 + legacyWindows,
    chromeExpired,
    tabletOsUnknown,
    lifecycle,
  };
}

function parseCyber(rows) {
  let totalControls = 0, naCount = 0, completeCount = 0, incompleteCount = 0;
  let cloudBackupStatus = "Unknown";
  let m365BackupStatus = "Unknown";
  let emailSecurityStatus = "Unknown";
  const controls = [];
  let headerIdx = -1;
  let areaIdx = 0;
  let statusIdx = 7;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    for (let c = 0; c < row.length; c++) {
      const k = normKey(row[c]);
      if (k === "control area" || (k.includes("control") && k.includes("area"))) {
        headerIdx = i;
        areaIdx = c;
        break;
      }
    }
    if (headerIdx >= 0) break;
  }

  if (headerIdx >= 0) {
    const header = rows[headerIdx] || [];
    for (let c = 0; c < header.length; c++) {
      const k = normKey(header[c]);
      if (k === "control area" || (k.includes("control") && k.includes("area"))) areaIdx = c;
      if (k === "status" || k.includes("status") || k.includes("readiness")) statusIdx = c;
    }
  }

  const start = headerIdx >= 0 ? headerIdx + 1 : 0;
  for (let i = start; i < rows.length; i++) {
    const row = rows[i] || [];
    const areaText = String(row[areaIdx] || "").trim();
    if (!areaText) continue;
    totalControls += 1;
    const area = areaText.toLowerCase();
    const status = `${row[statusIdx] || ""}`.trim() || "Unknown";
    const stage = classifyCyberStatus(status);
    if (stage === "na") naCount += 1;
    if (stage === "complete") completeCount += 1;
    else incompleteCount += 1;
    controls.push({ area: areaText, status, stage });

    if (area.includes("cloud backup")) {
      cloudBackupStatus = status || "Unknown";
    }
    if (area.includes("365 backup") || area.includes("m365 backup") || area.includes("microsoft 365 backup")) {
      m365BackupStatus = status;
    }
    if (area.includes("email security")) {
      emailSecurityStatus = status;
    }
  }
  return {
    totalControls,
    naCount,
    completeCount,
    incompleteCount,
    cloudBackupStatus,
    m365BackupStatus,
    emailSecurityStatus,
    controls,
  };
}

function parseMigration(rows) {
  let total = 0, ready = 0, inProgress = 0, remediation = 0;
  const items = [];
  const seen = new Set();
  let headerIdx = -1;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const hasArea = normKey(row[0]) === "area";
    const hasCheck = normKey(row[1]) === "check";
    if (hasArea && hasCheck) {
      headerIdx = i;
      break;
    }
  }

  const header = headerIdx >= 0 ? (rows[headerIdx] || []) : [];
  const titleRow = headerIdx > 0 ? (rows[headerIdx - 1] || []) : [];
  const blocks = detectMigrationBlocks(header, titleRow);

  const start = headerIdx >= 0 ? headerIdx + 1 : 0;
  for (let i = start; i < rows.length; i++) {
    const row = rows[i];
    for (const block of blocks) {
      const area = String(row[block.areaIdx] || "").trim();
      const check = String(row[block.checkIdx] || "").trim();
      if (!area && !check) continue;
      const sourcePlatform = String(row[block.sourceIdx] || "").trim();
      if (!isMicrosoftMigrationRow(area, check, sourcePlatform, block.kind)) continue;

      const rawStatus = String(row[block.statusIdx] || "").trim();
      if (!rawStatus) continue;
      const action = String(row[block.actionIdx] || "").trim();
      const itemKey = `${area}|${check}|${rawStatus}|${action}`;
      if (seen.has(itemKey)) continue;
      seen.add(itemKey);

      total += 1;
      const stage = getMigrationStage(rawStatus);
      items.push({
        area: area || "-",
        check: check || "-",
        status: rawStatus || stage,
        stage,
        action: action || (stage === "Ready" ? "No action required" : "Define remediation action"),
      });

      if (stage === "Ready") ready += 1;
      if (stage === "In progress") inProgress += 1;
      if (stage === "Remediation required") remediation += 1;
    }
  }

  const remaining = items.filter((item) => item.stage !== "Ready");
  return { total, ready, inProgress, remediation, items, remaining };
}

function detectMigrationBlocks(header, titleRow = []) {
  if (!header || !header.length) {
    return [{ areaIdx: 0, checkIdx: 1, sourceIdx: 2, statusIdx: 5, actionIdx: 8, kind: "microsoft" }];
  }

  const keys = header.map((h) => normKey(h));
  const areaStarts = [];
  for (let i = 0; i < keys.length; i++) {
    if (keys[i] === "area") areaStarts.push(i);
  }
  if (!areaStarts.length) {
    return [{ areaIdx: 0, checkIdx: 1, sourceIdx: 2, statusIdx: 5, actionIdx: 8, kind: "microsoft" }];
  }

  const blocks = [];
  for (let b = 0; b < areaStarts.length; b++) {
    const start = areaStarts[b];
    const end = b < areaStarts.length - 1 ? areaStarts[b + 1] : keys.length;
    const block = {
      areaIdx: start,
      checkIdx: Math.min(start + 1, Math.max(start, end - 1)),
      sourceIdx: Math.min(start + 2, Math.max(start, end - 1)),
      statusIdx: Math.min(start + 5, Math.max(start, end - 1)),
      actionIdx: Math.min(start + 8, Math.max(start, end - 1)),
      kind: "other",
    };

    for (let i = start; i < end; i++) {
      const key = keys[i];
      if (key === "check") block.checkIdx = i;
      if (key === "source platform" || key.includes("source platform")) block.sourceIdx = i;
      if (key === "status" || key.includes("status")) block.statusIdx = i;
      if (key === "remediation required") block.actionIdx = i;
      else if (key.includes("action") || key.includes("recommendation")) block.actionIdx = i;
    }
    const laneTitle = normKey(titleRow[start] || "");
    if (laneTitle.includes("google")) block.kind = "google";
    else if (laneTitle.includes("microsoft") || laneTitle.includes("office 365") || laneTitle.includes("m365")) block.kind = "microsoft";
    else if (b === 0) block.kind = "microsoft";
    blocks.push(block);
  }

  return blocks;
}

function isMicrosoftMigrationRow(areaText, checkText, sourcePlatform = "", blockKind = "other") {
  const combined = `${areaText || ""} ${checkText || ""}`.toLowerCase();
  const source = String(sourcePlatform || "").toLowerCase();
  if (!combined.trim()) return false;

  const hasGoogle =
    combined.includes("google") ||
    combined.includes("gmail") ||
    combined.includes("google workspace") ||
    combined.includes("g suite") ||
    combined.includes("chromebook");
  const sourceIsGoogle =
    source.includes("google") ||
    source.includes("gmail") ||
    source.includes("workspace") ||
    source.includes("g suite") ||
    source.includes("chromebook");
  if (hasGoogle || sourceIsGoogle || blockKind === "google") return false;

  const hasMicrosoftSignal =
    combined.includes("microsoft") ||
    combined.includes("m365") ||
    combined.includes("office 365") ||
    combined.includes("entra") ||
    combined.includes("azure") ||
    combined.includes("intune") ||
    combined.includes("sharepoint") ||
    combined.includes("onedrive") ||
    combined.includes("exchange") ||
    combined.includes("teams") ||
    combined.includes("windows") ||
    combined.includes("active directory") ||
    combined.includes("ad ");
  const sourceIsMicrosoft =
    source.includes("microsoft") ||
    source.includes("office 365") ||
    source.includes("m365") ||
    source.includes("entra") ||
    source.includes("azure") ||
    source.includes("intune");

  if (blockKind === "microsoft") return true;
  return sourceIsMicrosoft || hasMicrosoftSignal;
}

function parseCore(rows) {
  let currentLicence = "Unknown";
  let inFeatureTable = false;
  let featureRows = 0;
  let configuredUnknown = 0;
  const featureItems = [];
  for (const row of rows) {
    if (String(row[6] || "").toLowerCase() === "current licence") currentLicence = row[7] || "Unknown";
    if (row[0] === "Feature Area" && row[1] === "Feature / Capability") { inFeatureTable = true; continue; }
    if (!inFeatureTable || !row[0]) continue;
    featureRows += 1;
    const featureArea = String(row[0] || "").trim();
    const featureName = String(row[1] || "").trim();
    const configuredRaw = String(row[7] || "").trim();
    const usageRaw = String(row[9] || "").trim();
    featureItems.push({
      featureArea,
      featureName,
      configured: configuredRaw,
      usage: usageRaw,
    });
    const configured = String(row[7] || "").toLowerCase();
    const usage = String(row[9] || "").toLowerCase();
    if ((configured === "yes" || configured.includes("yes")) && usage.includes("assumed")) configuredUnknown += 1;
  }
  return { currentLicence, featureRows, configuredUnknown, featureItems };
}

function parseSoftware(rows) {
  if (!rows.length) return { entries: [] };

  let headerIdx = -1;
  let nameIdx = 0;
  let statusIdx = -1;
  let categoryIdx = -1;
  let vendorIdx = -1;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const keys = row.map((cell) => normKey(cell));
    const hasSoftwareHeader = keys.some((k) => k === "software" || k.includes("application") || k.includes("product") || k === "name");
    if (!hasSoftwareHeader) continue;
    headerIdx = i;
    for (let c = 0; c < keys.length; c++) {
      const key = keys[c];
      if (key === "software" || key === "application" || key === "product" || key === "name" || key.includes("software name")) nameIdx = c;
      if (key === "status" || key.includes("status") || key.includes("configured")) statusIdx = c;
      if (key.includes("category") || key.includes("type")) categoryIdx = c;
      if (key.includes("vendor") || key.includes("supplier")) vendorIdx = c;
    }
    break;
  }

  const start = headerIdx >= 0 ? headerIdx + 1 : 0;
  const entries = [];
  for (let i = start; i < rows.length; i++) {
    const row = rows[i] || [];
    const cells = row.map((cell) => String(cell || "").trim());
    if (!cells.some((cell) => cell !== "")) continue;

    let name = String(row[nameIdx] || "").trim();
    if (!name) {
      const candidate = cells.find((cell) => cell);
      name = candidate || "";
    }
    if (!name) continue;

    const keyName = normKey(name);
    if (keyName === "software" || keyName === "application" || keyName === "product") continue;

    entries.push({
      name,
      status: statusIdx >= 0 ? String(row[statusIdx] || "").trim() : "",
      category: categoryIdx >= 0 ? String(row[categoryIdx] || "").trim() : "",
      vendor: vendorIdx >= 0 ? String(row[vendorIdx] || "").trim() : "",
    });
  }

  return { entries };
}

function parseQuestions(rows) {
  if (!rows.length) return { total: 0, open: 0 };
  let open = 0;
  for (const row of rows) {
    const status = String(row[7] || "").toLowerCase();
    if (status === "open") open += 1;
  }
  return { total: rows.length - 1, open };
}

function renderKpis(dashboard, infra, network, client, cyber, migration, questions) {
  const msReadyValue = migration.total > 0
    ? `${((migration.ready / migration.total) * 100).toFixed(1)}%`
    : (dashboard.msReadyPct !== "-" ? dashboard.msReadyPct : "-");
  const msRemediationValue = migration.total > 0
    ? migration.remediation
    : (dashboard.msRemediation !== "-" ? dashboard.msRemediation : 0);
  const cyberPctFallback = cyber.totalControls > 0
    ? `${((cyber.completeCount / cyber.totalControls) * 100).toFixed(1)}%`
    : "-";
  const cyberPctValue = dashboard.cyberPct !== "-" ? dashboard.cyberPct : cyberPctFallback;
  const kpis = [
    ["Servers", infra.totalServers],
    ["SAN switches", infra.sanSwitchCount],
    ["SAN storage arrays", infra.sanStorageCount],
    ["Core switches", network.core],
    ["Edge switches", network.edge],
    ["Firewalls", network.firewalls],
    ["Wi-Fi APs", network.aps],
    ["Client devices (qty)", client.totalQuantity || dashboard.devices],
    ["Windows devices", client.windowsDevices],
    ["Chromebooks", client.chromebooks],
    ["Tablets", client.tabletDevices],
    ["Cyber complete", cyberPctValue],
    ["Microsoft ready", msReadyValue],
    ["MS remediation items", msRemediationValue],
    ["Win Server 2012", infra.ws2012],
    ["Server critical", infra.serverCritical],
    ["Physical monitor-only", infra.physicalMonitor],
    ["APs OOW", network.apsOutOfWarranty],
    ["Switches OOW", network.switchesOutOfWarranty],
    ["Firewalls OOW", network.firewallsOutOfWarranty],
    ["Old OS devices", client.oldOsTotal],
    ["Wi-Fi 5 / Wi-Fi 6", `${network.wifi5} / ${network.wifi6}`],
    ["Open questions", questions.open],
    ["Migration checks", migration.total],
    ["Migration remediation", migration.remediation],
  ];
  els.kpiGrid.innerHTML = kpis.map(([label, value]) => `
    <div class="metric">
      <div class="label">${escapeHtml(label)}</div>
      <div class="value">${escapeHtml(String(value))}</div>
    </div>
  `).join("");
}

function renderRisks(infra, cyber, migration, client, core) {
  const critical = [];
  const inFlight = [];
  const strategic = [];

  if (infra.ws2012 > 0) critical.push(`${infra.ws2012} server workloads still on Windows Server 2012 (urgent upgrade).`);
  if (String(cyber.cloudBackupStatus).toLowerCase() === "n/a") critical.push("Cloud backup is not implemented.");
  if (cyber.naCount > 0) inFlight.push(`${cyber.naCount} cyber controls are marked N/A (not implemented).`);
  if (migration.remediation > 0) inFlight.push(`${migration.remediation} Microsoft remediation items remain open.`);
  if (client.windows10 > 0) inFlight.push(`${client.windows10} devices on Windows 10 require replacement planning.`);
  if (infra.physicalMonitor > 0) strategic.push(`${infra.physicalMonitor} physical server items are monitor/planning items (not immediate replacement).`);
  if (core.configuredUnknown > 0) strategic.push(`${core.configuredUnknown} configured capabilities need usage confirmation for ROI tracking.`);
  strategic.push("Trust-wide check: verify interim 365 backup and email security controls across remaining schools before migration.");

  const stageColumn = (title, toneClass, items, emptyText) => `
    <div class="migration-stage-card ${toneClass}">
      <div class="migration-stage-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${items.length}</span>
      </div>
      ${
        items.length
          ? `<ul class="list migration-stage-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : `<p class="muted">${escapeHtml(emptyText)}</p>`
      }
    </div>
  `;

  els.riskList.innerHTML = `
    <div class="migration-overview">
      <div class="metric">
        <div class="label">Critical risks</div>
        <div class="value ${critical.length > 0 ? "danger" : "ok"}">${critical.length}</div>
      </div>
      <div class="metric">
        <div class="label">Priority remediation risks</div>
        <div class="value ${inFlight.length > 0 ? "warn" : "ok"}">${inFlight.length}</div>
      </div>
      <div class="metric">
        <div class="label">Strategic watchpoints</div>
        <div class="value">${strategic.length}</div>
      </div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${stageColumn("Critical", "stage-remediation", critical, "No critical risks currently flagged.")}
      ${stageColumn("Priority remediation", "stage-progress", inFlight, "No priority remediation risks currently flagged.")}
      ${stageColumn("Strategic watchpoints", "stage-other", strategic, "No strategic watchpoints currently flagged.")}
    </div>
  `;
}

function renderCriticalCounts(infra, network, client) {
  const merged = mergeLifecycleStores([infra.lifecycle, network.lifecycle, client.lifecycle]);
  const warrantyFlags = merged.flagged.warranty.length;
  const supportFlags = merged.flagged.support.length;
  const serviceLifeFlags = merged.flagged.serviceLife.length;
  const serverOsFlags = infra.ws2012 + infra.ws2016;
  const clientOsFlags = client.windows10 + client.legacyWindows;
  const osFlags = serverOsFlags + clientOsFlags;
  const totalFlags = warrantyFlags + supportFlags + serviceLifeFlags + osFlags;

  const hardwareList = [];
  if (warrantyFlags > 0) hardwareList.push(`End of warranty flagged items: ${warrantyFlags}`);
  if (supportFlags > 0) hardwareList.push(`End of support flagged items: ${supportFlags}`);
  if (serviceLifeFlags > 0) hardwareList.push(`End of service life flagged items: ${serviceLifeFlags}`);

  const networkList = [];
  if (network.apsOutOfWarranty > 0) networkList.push(`Wi-Fi APs out of warranty: ${network.apsOutOfWarranty}`);
  if (network.switchesOutOfWarranty > 0) networkList.push(`Switches out of warranty: ${network.switchesOutOfWarranty}`);
  if (network.firewallsOutOfWarranty > 0) networkList.push(`Firewalls out of warranty: ${network.firewallsOutOfWarranty}`);

  const osList = [];
  if (infra.ws2012 > 0) osList.push(`Windows Server 2012 workloads: ${infra.ws2012} (urgent)`);
  if (infra.ws2016 > 0) osList.push(`Windows Server 2016 workloads: ${infra.ws2016} (monitor/plan)`);
  if (client.windows10 > 0) osList.push(`Windows 10 client devices: ${client.windows10}`);
  if (client.legacyWindows > 0) osList.push(`Legacy Windows (7/8/8.1) devices: ${client.legacyWindows}`);

  const stageColumn = (title, toneClass, items, emptyText) => `
    <div class="migration-stage-card ${toneClass}">
      <div class="migration-stage-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${items.length}</span>
      </div>
      ${
        items.length
          ? `<ul class="list migration-stage-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : `<p class="muted">${escapeHtml(emptyText)}</p>`
      }
    </div>
  `;

  els.criticalCounts.innerHTML = `
    <div class="migration-overview">
      <div class="metric">
        <div class="label">Total critical roadmap flags</div>
        <div class="value ${totalFlags > 0 ? "danger" : "ok"}">${totalFlags}</div>
      </div>
      <div class="metric">
        <div class="label">End of warranty flags</div>
        <div class="value ${warrantyFlags > 0 ? "warn" : "ok"}">${warrantyFlags}</div>
      </div>
      <div class="metric">
        <div class="label">End of support flags</div>
        <div class="value ${supportFlags > 0 ? "danger" : "ok"}">${supportFlags}</div>
      </div>
      <div class="metric">
        <div class="label">End of service life flags</div>
        <div class="value ${serviceLifeFlags > 0 ? "warn" : "ok"}">${serviceLifeFlags}</div>
      </div>
      <div class="metric">
        <div class="label">OS-based urgent/near-term flags</div>
        <div class="value ${osFlags > 0 ? "danger" : "ok"}">${osFlags}</div>
      </div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${stageColumn("Lifecycle milestones", "stage-progress", hardwareList, "No lifecycle milestone flags.")}
      ${stageColumn("Network & Wi-Fi", "stage-other", networkList, "No network warranty flags.")}
      ${stageColumn("Server and client OS", "stage-remediation", osList, "No OS risk flags.")}
    </div>
  `;
}

function renderMigration(migration) {
  if (!els.migrationSummary || !els.barReady) return;
  const readyPct = migration.total ? (migration.ready / migration.total) * 100 : 0;
  els.migrationSummary.innerHTML = `
    <p><strong>Total checks:</strong> ${migration.total}</p>
    <p><strong>Ready:</strong> ${migration.ready} | <strong>In progress:</strong> ${migration.inProgress} | <strong>Remediation required:</strong> ${migration.remediation}</p>
    <p class="muted">This dashboard prioritises Microsoft readiness for this school.</p>
  `;
  els.barReady.style.width = `${Math.max(0, Math.min(100, readyPct))}%`;
}

function renderMigrationActions(migration) {
  const outstanding = Math.max(0, migration.total - migration.ready);
  const all = (migration.items || []).slice(0, 200);
  const readyItems = all.filter((item) => item.stage === "Ready");
  const inProgressItems = all.filter((item) => item.stage === "In progress");
  const remediationItems = all.filter((item) => item.stage === "Remediation required");
  const otherItems = all.filter((item) => item.stage === "Other not ready");

  const renderStageColumn = (title, toneClass, items) => `
    <div class="migration-stage-card ${toneClass}">
      <div class="migration-stage-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${items.length}</span>
      </div>
      ${
        items.length
          ? `<ul class="list migration-stage-list">${items.map((item) => `
              <li>
                <strong>${escapeHtml(item.area)}:</strong> ${escapeHtml(item.check)}<br>
                ${escapeHtml(item.action)}
              </li>
            `).join("")}</ul>`
          : `<p class="muted">No items in this stage.</p>`
      }
    </div>
  `;

  els.migrationActions.innerHTML = `
    <div class="migration-overview">
      <div class="metric">
        <div class="label">Ready</div>
        <div class="value">${migration.ready}/${migration.total} (${migration.total ? ((migration.ready / migration.total) * 100).toFixed(1) : "0.0"}%)</div>
      </div>
      <div class="metric">
        <div class="label">In progress</div>
        <div class="value warn">${migration.inProgress}</div>
      </div>
      <div class="metric">
        <div class="label">Remediation required</div>
        <div class="value ${migration.remediation > 0 ? "danger" : "ok"}">${migration.remediation}</div>
      </div>
      <div class="metric">
        <div class="label">Total outstanding actions</div>
        <div class="value ${outstanding > 0 ? "warn" : "ok"}">${outstanding}</div>
      </div>
    </div>
    <div class="migration-board">
      ${renderStageColumn("Ready", "stage-ready", readyItems)}
      ${renderStageColumn("In progress", "stage-progress", inProgressItems)}
      ${renderStageColumn("Remediation required", "stage-remediation", remediationItems)}
      ${renderStageColumn("Other not ready", "stage-other", otherItems)}
    </div>
    <p class="muted">Board gate: close critical remediation before migration go/no-go.</p>
  `;
}

function getMigrationStage(statusText) {
  const s = normKey(statusText);
  if (s === "ready" || s === "completed" || s === "complete") return "Ready";
  if (s.includes("in progress")) return "In progress";
  if (s.includes("remediation")) return "Remediation required";
  if (s.includes("not ready")) return "Other not ready";
  return "Other not ready";
}

function isCyberGap(statusText) {
  const s = normKey(statusText);
  if (!s || s === "unknown") return true;
  if (s === "n/a") return true;
  if (s.includes("not implemented")) return true;
  if (s.includes("required")) return true;
  if (s.includes("missing")) return true;
  return false;
}

function classifyCyberStatus(statusText) {
  const s = normKey(statusText);
  if (!s || s === "unknown" || s === "tbc" || s === "to be confirmed") return "incomplete";
  if (s === "n/a" || s === "na") return "na";
  if (s.includes("not implemented") || s.includes("not complete") || s.includes("not configured")) return "incomplete";
  if (s.includes("remediation")) return "incomplete";
  if (s.includes("in progress") || s.includes("partially") || s.includes("partial")) return "incomplete";
  if (s === "no" || s.includes("missing") || s.includes("required")) return "incomplete";
  if (s === "yes" || s.includes("complete") || s.includes("implemented") || s.includes("configured")) return "complete";
  return "incomplete";
}

function mapBasicsStatusLabel(statusText) {
  const s = normKey(statusText);
  if (!s || s === "unknown" || s === "tbc" || s === "to be confirmed") return "Unknown";
  if (s.includes("in progress") || s.includes("partial") || s.includes("partially") || s.includes("remediation")) return "Partial";
  if (s === "n/a" || s === "na" || s === "no") return "Not In Place";
  if (s.includes("not implemented") || s.includes("not complete") || s.includes("not configured") || s.includes("missing") || s.includes("required")) return "Not In Place";
  if (s === "yes" || s.includes("complete") || s.includes("implemented") || s.includes("configured") || s.includes("enabled") || s.includes("active")) return "In Place";
  return "Unknown";
}

function mapCoreFeatureStatus(feature) {
  const configured = normKey(feature.configured);
  const usage = normKey(feature.usage);
  if (
    (configured.includes("yes") || configured.includes("enabled") || configured.includes("configured") || configured.includes("implemented")) &&
    (usage.includes("yes") || usage.includes("used") || usage.includes("active") || usage.includes("adopted") || usage.includes("complete"))
  ) return "In Place";
  if (configured.includes("yes") || configured.includes("enabled") || configured.includes("configured") || configured.includes("implemented")) return "Partial";
  if (configured === "no" || configured.includes("not")) return "Not In Place";
  if (usage.includes("partial") || usage.includes("in progress") || usage.includes("assumed")) return "Partial";
  if (usage === "no" || usage.includes("not used")) return "Not In Place";
  return "Unknown";
}

function mapSoftwareEvidenceStatus(entry) {
  const fromStatus = mapBasicsStatusLabel(entry.status);
  if (fromStatus !== "Unknown") return fromStatus;
  return "Partial";
}

function aggregateBasicsLabels(labels) {
  const clean = (labels || []).filter((label) => !!label);
  if (!clean.length) return "Unknown";
  if (clean.includes("In Place")) return "In Place";
  if (clean.includes("Partial")) return "Partial";
  if (clean.includes("Not In Place")) return "Not In Place";
  return "Unknown";
}

function basicsStatusClassName(status) {
  if (status === "In Place") return "status-in-place";
  if (status === "Partial") return "status-partial";
  if (status === "Not In Place") return "status-not-in-place";
  return "status-unknown";
}

function formatBasicsEvidence(evidenceItems, fallback) {
  if (!evidenceItems.length) return fallback;
  const preview = evidenceItems.slice(0, 2).map((item) => item.note).join(" | ");
  const extra = evidenceItems.length > 2 ? ` (+${evidenceItems.length - 2} more)` : "";
  return `${preview}${extra}`;
}

function findControlsByKeywords(controls, keywords) {
  const keys = (keywords || []).map((k) => normKey(k)).filter((k) => k);
  return (controls || []).filter((control) => {
    const area = normKey(control.area);
    return keys.some((key) => area.includes(key));
  });
}

function findCoreFeaturesByKeywords(features, keywords) {
  const keys = (keywords || []).map((k) => normKey(k)).filter((k) => k);
  return (features || []).filter((feature) => {
    const combined = normKey(`${feature.featureArea || ""} ${feature.featureName || ""}`);
    return keys.some((key) => combined.includes(key));
  });
}

function findSoftwareByKeywords(entries, keywords) {
  const keys = (keywords || []).map((k) => normKey(k)).filter((k) => k);
  return (entries || []).filter((entry) => {
    const combined = normKey(`${entry.name || ""} ${entry.status || ""} ${entry.category || ""} ${entry.vendor || ""}`);
    return keys.some((key) => combined.includes(key));
  });
}

function buildBrilliantBasicsRows(cyber, core, software) {
  const controls = cyber.controls || [];
  const features = core.featureItems || [];
  const softwareEntries = software.entries || [];

  function capabilityFromEvidence(capability, evidenceItems, fallbackPrompt) {
    return {
      capability,
      status: aggregateBasicsLabels(evidenceItems.map((item) => item.status)),
      evidence: formatBasicsEvidence(evidenceItems, fallbackPrompt),
    };
  }

  const m365Controls = findControlsByKeywords(controls, ["365 backup", "m365 backup", "microsoft 365 backup"]);
  const emailControls = findControlsByKeywords(controls, ["email security"]);
  const mfaControls = findControlsByKeywords(controls, ["mfa", "multi-factor", "multifactor", "conditional access"]);
  const mfaFeatures = findCoreFeaturesByKeywords(features, ["mfa", "multi-factor", "multifactor", "conditional access", "entra"]);
  const antivirusControls = findControlsByKeywords(controls, ["anti-virus", "antivirus", "endpoint protection", "defender", "sophos"]);
  const antivirusSoftware = findSoftwareByKeywords(softwareEntries, ["anti-virus", "antivirus", "defender", "sophos", "endpoint protection", "edr"]);
  const antivirusFeatures = findCoreFeaturesByKeywords(features, ["defender", "endpoint protection", "anti-virus", "antivirus", "edr"]);
  const rmmControls = findControlsByKeywords(controls, ["rmm", "remote monitoring", "remote management", "remote monitoring and management"]);
  const rmmSoftware = findSoftwareByKeywords(softwareEntries, ["rmm", "remote monitoring", "remote management", "remote monitoring and management"]);
  const rmmFeatures = findCoreFeaturesByKeywords(features, ["rmm", "remote monitoring", "remote management"]);

  const m365Evidence = [
    { status: mapBasicsStatusLabel(cyber.m365BackupStatus || "Unknown"), note: `Parsed field m365BackupStatus: ${cyber.m365BackupStatus || "Unknown"}` },
    ...m365Controls.map((control) => ({ status: mapBasicsStatusLabel(control.status), note: `${control.area}: ${control.status || "Unknown"}` })),
  ];

  const emailEvidence = [
    { status: mapBasicsStatusLabel(cyber.emailSecurityStatus || "Unknown"), note: `Parsed field emailSecurityStatus: ${cyber.emailSecurityStatus || "Unknown"}` },
    ...emailControls.map((control) => ({ status: mapBasicsStatusLabel(control.status), note: `${control.area}: ${control.status || "Unknown"}` })),
  ];

  const mfaEvidence = [
    ...mfaControls.map((control) => ({ status: mapBasicsStatusLabel(control.status), note: `${control.area}: ${control.status || "Unknown"}` })),
    ...mfaFeatures.map((feature) => ({
      status: mapCoreFeatureStatus(feature),
      note: `${feature.featureName || feature.featureArea || "Core feature"}: Configured=${feature.configured || "Unknown"}, Usage=${feature.usage || "Unknown"}`,
    })),
  ];

  const antivirusEvidence = [
    ...antivirusControls.map((control) => ({ status: mapBasicsStatusLabel(control.status), note: `${control.area}: ${control.status || "Unknown"}` })),
    ...antivirusSoftware.map((entry) => ({
      status: mapSoftwareEvidenceStatus(entry),
      note: `${entry.name}${entry.status ? `: ${entry.status}` : " (listed in software inventory)"}`,
    })),
    ...antivirusFeatures.map((feature) => ({
      status: mapCoreFeatureStatus(feature),
      note: `${feature.featureName || feature.featureArea || "Core feature"}: Configured=${feature.configured || "Unknown"}, Usage=${feature.usage || "Unknown"}`,
    })),
  ];

  const rmmEvidence = [
    ...rmmControls.map((control) => ({ status: mapBasicsStatusLabel(control.status), note: `${control.area}: ${control.status || "Unknown"}` })),
    ...rmmSoftware.map((entry) => ({
      status: mapSoftwareEvidenceStatus(entry),
      note: `${entry.name}${entry.status ? `: ${entry.status}` : " (listed in software inventory)"}`,
    })),
    ...rmmFeatures.map((feature) => ({
      status: mapCoreFeatureStatus(feature),
      note: `${feature.featureName || feature.featureArea || "Core feature"}: Configured=${feature.configured || "Unknown"}, Usage=${feature.usage || "Unknown"}`,
    })),
  ];

  return [
    capabilityFromEvidence("Microsoft 365 backup", m365Evidence, "No explicit Microsoft 365 backup evidence found — validate in workbook."),
    capabilityFromEvidence("Email security", emailEvidence, "No explicit email security evidence found — validate in workbook."),
    capabilityFromEvidence("MFA / Conditional Access", mfaEvidence, "No explicit MFA / Conditional Access evidence found — validate in workbook."),
    capabilityFromEvidence("Anti-virus", antivirusEvidence, "No explicit anti-virus evidence found — validate in workbook."),
    capabilityFromEvidence("RMM", rmmEvidence, "No explicit RMM evidence found in parsed controls/software/core features — validate in workbook."),
  ];
}

function createLifecycleStore() {
  return {
    flagged: { warranty: [], support: [], serviceLife: [] },
    good: { warranty: [], support: [], serviceLife: [] },
  };
}

function addLifecycleEntry(store, milestone, statusObj, itemLabel) {
  if (statusObj && statusObj.include === false) return;
  const entry = `${itemLabel} - ${statusObj.reason}`;
  if (statusObj.flagged) store.flagged[milestone].push(entry);
  else store.good[milestone].push(entry);
}

function lifecycleStatus(kind, statusRaw, dateRaw, parsedDate) {
  const statusKey = normKey(statusRaw);
  const rawKey = normKey(dateRaw);
  const rawIsNa = isNaToken(dateRaw);
  const prettyDate = parsedDate ? formatDate(parsedDate) : "No date";

  if (kind === "warranty" && statusKey.includes("expired")) {
    return { flagged: true, reason: "Warranty status: Expired" };
  }
  if (rawIsNa) {
    return { flagged: true, reason: "Date marked N/A (treated as expired)" };
  }
  if (isPastDate(parsedDate)) {
    return { flagged: true, reason: `Date passed (${prettyDate})` };
  }
  if (kind === "warranty" && (statusKey.includes("active") || statusKey.includes("in warranty") || statusKey.includes("vendor support"))) {
    return { flagged: false, reason: `Status: ${statusRaw || "Active"}` };
  }
  if (parsedDate) {
    return { flagged: false, reason: `Date valid until ${prettyDate}` };
  }
  if (!rawKey) {
    return { flagged: false, reason: "No lifecycle data recorded", include: false };
  }
  return { flagged: false, reason: `Status: ${statusRaw || "Not flagged"}` };
}

function isUnsupportedServerOs(osText) {
  const os = normKey(osText);
  return (
    os.includes("2003") ||
    os.includes("2008") ||
    os.includes("2012")
  );
}

function isNaToken(value) {
  const s = normKey(value);
  return s === "n/a" || s === "na";
}

function formatDate(value) {
  if (!(value instanceof Date) || isNaN(value)) return "";
  return value.toLocaleDateString("en-GB");
}

function renderLifecycleDetails(lifecycle) {
  if (!lifecycle) return "";
  return `
    <details class="details-list lifecycle-dropdown" open>
      <summary><strong>Items to flag</strong></summary>
      ${renderLifecycleColumns(lifecycle.flagged, "No flagged items in this milestone.")}
    </details>
  `;
}

function renderLifecycleColumns(group, emptyText) {
  return `
    <div class="lifecycle-columns">
      ${renderLifecycleColumn("End of warranty", group.warranty || [], emptyText)}
      ${renderLifecycleColumn("End of support", group.support || [], emptyText)}
      ${renderLifecycleColumn("End of service life", group.serviceLife || [], emptyText)}
    </div>
  `;
}

function renderLifecycleColumn(title, items, emptyText) {
  return `
    <div class="lifecycle-col">
      <div class="lifecycle-col-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${items.length}</span>
      </div>
      ${
        items.length
          ? `<ul class="list lifecycle-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : `<p class="muted">${escapeHtml(emptyText)}</p>`
      }
    </div>
  `;
}

function renderLifecycleSummary(infra, network, client) {
  const merged = mergeLifecycleStores([infra.lifecycle, network.lifecycle, client.lifecycle]);
  const flaggedWarrantyItems = merged.flagged.warranty;
  const flaggedSupportItems = merged.flagged.support;
  const flaggedServiceLifeItems = merged.flagged.serviceLife;

  const endOfWarranty = flaggedWarrantyItems.length;
  const endOfServiceLife = flaggedServiceLifeItems.length;
  const endOfSupport = flaggedSupportItems.length;

  const stageColumn = (title, toneClass, items) => `
    <div class="migration-stage-card ${toneClass}">
      <div class="migration-stage-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${items.length}</span>
      </div>
      ${
        items.length
          ? `<ul class="list migration-stage-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : `<p class="muted">No flagged items in this milestone.</p>`
      }
    </div>
  `;

  els.lifecycleSummary.innerHTML = `
    <div class="migration-overview">
      <div class="metric">
        <div class="label">End of warranty</div>
        <div class="value ${endOfWarranty > 0 ? "warn" : "ok"}">${endOfWarranty}</div>
      </div>
      <div class="metric">
        <div class="label">End of service life</div>
        <div class="value ${endOfServiceLife > 0 ? "warn" : "ok"}">${endOfServiceLife}</div>
      </div>
      <div class="metric">
        <div class="label">End of support (critical)</div>
        <div class="value ${endOfSupport > 0 ? "danger" : "ok"}">${endOfSupport}</div>
      </div>
      <div class="metric">
        <div class="label">Old OS devices</div>
        <div class="value ${client.oldOsTotal > 0 ? "danger" : "ok"}">${client.oldOsTotal}</div>
      </div>
    </div>
    <div class="migration-board lifecycle-status-board">
      ${stageColumn("End of warranty (flagged)", "stage-progress", flaggedWarrantyItems)}
      ${stageColumn("End of support (flagged)", "stage-remediation", flaggedSupportItems)}
      ${stageColumn("End of service life (flagged)", "stage-other", flaggedServiceLifeItems)}
    </div>
    <p class="muted">End of support is the final critical inflection point for replacement.</p>
  `;
}

function mergeLifecycleStores(stores) {
  const merged = createLifecycleStore();
  for (const store of stores) {
    if (!store) continue;
    for (const key of ["warranty", "support", "serviceLife"]) {
      merged.flagged[key] = merged.flagged[key].concat(store.flagged?.[key] || []);
      merged.good[key] = merged.good[key].concat(store.good?.[key] || []);
    }
  }
  return merged;
}

function populateSheetSelect() {
  els.sheetSelect.innerHTML = workbook.SheetNames.map((name) => `
    <option value="${escapeAttr(name)}">${escapeHtml(name)}</option>
  `).join("");
}

function renderSheetTable() {
  if (!workbook) return;
  const sheetName = els.sheetSelect.value || workbook.SheetNames[0];
  const rows = getRows(sheetName);
  const nonEmptyRows = rows.filter((r) => r.some((c) => c !== null && c !== undefined && String(c).trim() !== ""));

  const maxCols = nonEmptyRows.reduce((m, r) => Math.max(m, r.length), 0);
  const header = Array.from({ length: maxCols }, (_, i) => `Col ${i + 1}`);
  const limited = nonEmptyRows.slice(0, 1000);

  const thead = els.sheetTable.querySelector("thead");
  const tbody = els.sheetTable.querySelector("tbody");
  thead.innerHTML = `<tr>${header.map((h) => `<th>${h}</th>`).join("")}</tr>`;
  tbody.innerHTML = limited.map((row) => `<tr>${header.map((_, i) => `<td>${escapeHtml(formatCell(row[i]))}</td>`).join("")}</tr>`).join("");
  els.sheetInfo.textContent = `${sheetName}: showing ${limited.length} rows of ${nonEmptyRows.length} non-empty rows.`;
}

function getRows(sheetName) {
  if (!workbook || !workbook.Sheets[sheetName]) return [];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: false });
}

function pct(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "-";
  if (raw.endsWith("%")) {
    const n = Number(raw.replace("%", "").trim());
    return Number.isFinite(n) ? `${n.toFixed(1)}%` : "-";
  }
  const n = Number(raw);
  if (!Number.isFinite(n)) return "-";
  if (n > 1) return `${n.toFixed(1)}%`;
  return `${(n * 100).toFixed(1)}%`;
}

function parseExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) return value;
  const parsed = new Date(value);
  if (!isNaN(parsed)) return parsed;
  return null;
}

function isPastDate(value) {
  return !!(value && value instanceof Date && value < new Date());
}

function normKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isOutOfWarranty(status, expiryDate) {
  const s = String(status || "").toLowerCase();
  if (s.includes("expired")) return true;
  if (s.includes("to be found") || s.includes("?")) return false;
  if (s === "n/a" || s === "active" || s === "in warranty" || s === "vendor support") return false;
  if (expiryDate && expiryDate < new Date()) return true;
  return false;
}

function formatCell(v) {
  if (v === null || v === undefined || v === "") return "";
  return String(v);
}

function isSectionHeaderRow(row) {
  if (!row || !String(row[0] || "").trim()) return false;
  const hasOnlyFirstCell = row.slice(1).every((cell) => String(cell || "").trim() === "");
  return hasOnlyFirstCell;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttr(text) {
  return escapeHtml(text).replaceAll("\n", " ");
}

function exportWord() {
  if (!workbook) return;
  const html = buildExportHtml("word");
  const blob = new Blob(["\ufeff", html], { type: "application/msword" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "audit-dashboard-export.doc";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}

function exportWeb() {
  if (!workbook) return;
  const html = buildExportHtml("web");
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = "audit-report-visual.html";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  if (els.reportLink) {
    const safeUrl = escapeAttr(objectUrl);
    els.reportLink.innerHTML = `
      Web report link: <a href="${safeUrl}" target="_blank" rel="noopener noreferrer">Open visual report</a>
      <button id="copy-report-link" type="button">Copy link</button>
    `;
    const copyBtn = document.getElementById("copy-report-link");
    if (copyBtn && navigator.clipboard) {
      copyBtn.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(objectUrl);
          copyBtn.textContent = "Copied";
        } catch (_) {
          copyBtn.textContent = "Copy failed";
        }
      }, { once: true });
    }
  }
}

async function exportPdf() {
  if (!workbook || !latestReport) return;
  await ensurePdfLibrariesLoaded();
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });

  const data = latestReport;
  const { infra, network, client, cyber, migration, core, dashboard } = data;
  const software = data.software || { entries: [] };
  const lifecycle = mergeLifecycleStores([infra.lifecycle, network.lifecycle, client.lifecycle]);
  const endWarranty = (lifecycle.flagged.warranty || []).length;
  const endSupport = (lifecycle.flagged.support || []).length;
  const endServiceLife = (lifecycle.flagged.serviceLife || []).length;
  const msReadyPct = migration.total ? `${((migration.ready / migration.total) * 100).toFixed(1)}%` : "-";
  const cyberPctFallback = cyber.totalControls > 0 ? `${((cyber.completeCount / cyber.totalControls) * 100).toFixed(1)}%` : "-";
  const cyberPct = dashboard.cyberPct !== "-" ? dashboard.cyberPct : cyberPctFallback;

  const infraStatus = (infra.ws2012 > 0 || infra.serverCritical > 0) ? "Red" : "Amber";
  const networkStatus = (network.apsOutOfWarranty + network.switchesOutOfWarranty + network.firewallsOutOfWarranty) > 0 ? "Red" : "Amber";
  const clientStatus = client.oldOsTotal > 0 ? "Amber" : "Green";
  const cyberStatus = cyber.incompleteCount > 0 ? "Amber / Red" : "Green";
  const tenancyStatus = migration.remediation > 0 ? "Red" : "Amber";
  const coreStatus = core.configuredUnknown > 0 ? "Amber" : "Green";
  const brilliantBasicsRows = buildBrilliantBasicsRows(cyber, core, software);

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  doc.setFillColor(10, 48, 88);
  doc.rect(0, 0, pageWidth, pageHeight, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(34);
  const coverTitleLines = doc.splitTextToSize("IT Audit Summary, Findings and Roadmap", pageWidth - 80);
  const coverTitleY = 120;
  doc.text(coverTitleLines, 40, coverTitleY);
  const coverTitleBottomY = coverTitleY + ((coverTitleLines.length - 1) * 38);
  doc.setTextColor(125, 211, 252);
  doc.setFontSize(24);
  const schoolY = coverTitleBottomY + 44;
  doc.text("Sir James Smith's School", 40, schoolY);
  doc.setTextColor(219, 234, 254);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(14);
  const intro = "Thank you for undertaking the Unleashed site audit and readiness assessment. This report summarises the current IT estate, key findings, principal risks and a prioritised improvement roadmap.";
  const introY = schoolY + 50;
  const introLines = doc.splitTextToSize(intro, pageWidth - 80);
  doc.text(introLines, 40, introY);
  doc.setTextColor(147, 197, 253);
  doc.setFont("helvetica", "bold");
  const focusY = introY + (introLines.length * 16) + 28;
  doc.text("Audit focus: Infrastructure • Cyber Security • Migration Readiness", 40, focusY);

  const chapters = [
    {
      title: "Executive Summary",
      intro: "This report translates the audit workbook into a board-level narrative with decisions, risks and a practical improvement roadmap.",
      head: ["Headline area", "Position", "Summary"],
      body: [
        ["Infrastructure lifecycle", infraStatus, `${infra.totalServers} servers recorded. Support-end and lifecycle flags require managed refresh planning.`],
        ["Cyber assurance", cyberStatus, `Cyber completion ${cyberPct}; not complete controls ${cyber.incompleteCount}.`],
        ["Migration readiness (Microsoft)", tenancyStatus, `${msReadyPct} ready, ${migration.remediation} remediation items remain.`],
        ["Core application / A3", coreStatus, `Configured-but-usage-unknown capabilities: ${core.configuredUnknown}.`],
      ],
    },
    {
      title: "RAG Summary",
      intro: "The RAG view provides an executive risk lens across each audited domain so leadership can prioritise remediation sequencing and governance ownership.",
      head: ["Area", "Status", "Comment"],
      body: [
        ["Data Infrastructure", infraStatus, "Server lifecycle and OS risks require phased action."],
        ["Network and Wi-Fi", networkStatus, "Warranty and lifecycle flags across switching, APs and firewall estate."],
        ["Client Compute", clientStatus, "Unsupported/ageing OS cohorts require replacement planning."],
        ["Cyber Security", cyberStatus, "Key controls still not complete; cloud backup and assurance remain priorities."],
        ["Tenancy Readiness", tenancyStatus, "Remediation closure needed before migration go/no-go."],
        ["Core Application / A3", coreStatus, "A3 optimisation and feature utilisation review remains open."],
      ],
    },
    {
      title: "Lifecycle and Critical Milestones",
      intro: "Lifecycle milestones should be treated as planning controls: warranty is flexible where service and support remain valid, but end of support is the critical replacement trigger.",
      head: ["Milestone", "Flagged count", "Planning interpretation"],
      body: [
        ["End of warranty", String(endWarranty), "Can be extension candidates where support/service life still valid."],
        ["End of support (critical)", String(endSupport), "Critical replacement/upgrade trigger."],
        ["End of service life", String(endServiceLife), "Higher risk of reactive failure and diminishing support options."],
      ],
    },
    {
      title: "Migration Readiness – Actions",
      intro: "Only remediation-required migration actions are listed below so this section can operate as a practical cutover control register with clear immediate actions.",
      head: ["Area", "Status", "Main risk / gap", "Immediate action"],
      body: (migration.items || [])
        .filter((item) => item.stage === "Remediation required")
        .slice(0, 25)
        .map((item) => [item.area, item.stage, item.check, item.action]),
    },
    {
      title: "Data Infrastructure – Servers, SAN and Storage",
      intro: "The data infrastructure position should be managed against the agreed cloud-first approach: retire or migrate workloads where viable, and replace only where local hosting remains essential.",
      head: ["Infrastructure item", "Position", "Main concern", "Recommended response"],
      body: [
        ["Physical server estate", `${infra.totalServers} records`, `${infra.ws2012} unsupported OS workloads and ${infra.serverCritical} immediate candidates.`, "Prioritise unsupported OS and support-end cohorts; phase cloud-first transition."],
        ["SAN / storage", `${infra.sanSwitchCount} SAN switches, ${infra.sanStorageCount} arrays`, "Operational dependency and lifecycle validation required.", "Validate support, utilisation and target-state alignment."],
        ["Virtual machines", `${infra.virtual} virtual workloads`, "Guest OS risk is key (host lifecycle inherited from physical).", "Treat unsupported VM OS as urgent remediation."],
      ],
    },
    {
      title: "Network and Wi-Fi – Refresh Direction",
      intro: "Network and wireless refresh should be coordinated as one programme to reach Wi-Fi 7 readiness, including 10Gb-capable core, 2.5Gb edge and modern AP coverage.",
      head: ["Network area", "Position", "Main concern", "Recommended response"],
      body: [
        ["Core/edge switching", `Core ${network.core}, Edge ${network.edge}`, "Lifecycle and warranty pressure across switching estate.", "Phase refresh with core-first then edge modernisation."],
        ["Wi-Fi platform", `${network.aps} APs`, "Mix of Wi-Fi 5/6 and warranty exposure.", "Design Wi-Fi 7 journey with switching/PoE alignment."],
        ["Firewalls", `${network.firewalls}`, "Supportability and replacement timing must be controlled.", "Maintain supported edge security path."],
      ],
    },
    {
      title: "Client Compute – Estate and Lifecycle",
      intro: "For client compute, the primary risk driver is unsupported operating systems. Devices on unsupported Windows versions should be treated as priority replacement candidates.",
      head: ["Device area", "Position", "Main concern", "Recommended response"],
      body: [
        ["Windows", `${client.windowsDevices}`, `Windows 10 / legacy cohorts: ${client.oldOsTotal} devices.`, "Plan staged replacement and supported OS compliance."],
        ["Chromebooks", `${client.chromebooks}`, `Expired ChromeOS cohorts: ${client.chromeExpired}.`, "Validate AUE lifecycle and replacement priority."],
        ["Tablet / Mac / other", `${client.tabletDevices + client.macDevices + client.otherDevices}`, "Consistency and governance quality vary by device class.", "Improve ownership, OS visibility and policy standards."],
      ],
    },
    {
      title: "Cyber Security Findings",
      intro: "Cyber remediation should prioritise controls marked incomplete or N/A, with cloud backup and documented recovery assurance treated as immediate resilience requirements.",
      head: ["Cyber area", "Position", "Main concern", "Recommended response"],
      body: [
        ["Cyber posture", cyberPct, `${cyber.incompleteCount} controls not complete; ${cyber.naCount} marked N/A.`, "Close N/A/incomplete controls with clear ownership and evidence."],
        ["Microsoft 365 backup", String(cyber.m365BackupStatus), "Interim controls may be required pre-tenancy migration.", "Confirm control ownership and delivery timeline."],
        ["Email security / Cloud backup", `${cyber.emailSecurityStatus} / ${cyber.cloudBackupStatus}`, "Cloud backup remains a critical resilience control.", "Prioritise cloud backup and email security assurance."],
      ],
    },
    {
      title: "Core Application and A3 Optimisation",
      intro: "This section focuses on value realisation from current licensing: enabled features should be evidenced as actively used, and non-enabled capability should be treated as a missed opportunity.",
      head: ["Area", "Position", "Main concern", "Recommended response"],
      body: [
        ["Current licence", String(core.currentLicence), "Feature enablement may exceed proven operational usage.", "Run usage validation and optimise A3 baseline."],
        ["Configured-but-usage-unknown", String(core.configuredUnknown), "Potential missed value and governance gap.", "Map enabled features to active operational use."],
      ],
    },
    {
      title: "Final Recommendations by Workbook Tab",
      intro: "These recommendations provide a phased execution plan that aligns lifecycle pressure, remediation urgency and migration readiness into a single delivery timeline.",
      head: ["Period", "Area", "Specific actions", "Timing driver"],
      body: [
        ["Immediate", "Migration readiness", "Close remediation actions and complete evidence pack for go/no-go.", `${migration.remediation} remediation items open.`],
        ["0-3 months", "Cyber security", "Prioritise incomplete/N/A controls; evidence 365 backup, email security, cloud backup.", `${cyber.incompleteCount} not complete controls.`],
        ["0-6 months", "Data & network lifecycle", "Address support-end/service-end flags and validate extension candidates.", `${endSupport + endServiceLife} support/service-life flags.`],
        ["6-12 months", "Client compute", "Reduce unsupported OS cohorts and lifecycle exception devices.", `${client.oldOsTotal} old OS devices.`],
      ],
    },
  ];

  const pdfExpansionByChapter = {
    4: {
      currentState: [
        `Microsoft readiness currently ${msReadyPct}.`,
        `${migration.remediation} remediation-required migration checks remain open.`,
        `${migration.inProgress} actions are in progress and need owner follow-through.`,
      ],
      whatThisMeans: "Cutover confidence depends on closing remediation items and validating dependencies across identity, mail, collaboration and applications.",
      priorityActions: {
        now: ["Close remediation-required rows and confirm evidence.", "Validate DNS, mail flow and identity controls."],
        d90: ["Complete UAT and publish go/no-go decision pack."],
        y12: ["Transition to post-migration optimisation and decommission governance."],
      },
    },
    5: {
      currentState: [
        `${infra.totalServers} records with ${infra.ws2012} legacy workloads.`,
        `${infra.serverCritical} immediate server candidates identified.`,
        `${infra.sanSwitchCount} SAN switches and ${infra.sanStorageCount} storage arrays require validation.`,
      ],
      whatThisMeans: "Server and storage strategy should prioritise unsupported workloads and reduce local dependency where cloud alternatives are viable.",
      priorityActions: {
        now: ["Remediate unsupported server OS workloads.", "Confirm SAN/storage support state and dependency mapping."],
        d90: ["Agree retain vs migrate vs replace decisions."],
        y12: ["Execute phased refresh only for services that must remain local."],
      },
    },
    6: {
      currentState: [
        `Core ${network.core} and edge ${network.edge} switching in mixed lifecycle state.`,
        `${network.aps} APs with Wi-Fi 7 readiness gap.`,
        `${network.firewalls} firewall records requiring support-path validation.`,
      ],
      whatThisMeans: "Wi-Fi 7 readiness requires coordinated core, edge and AP modernisation rather than isolated hardware swaps.",
      priorityActions: {
        now: ["Validate network inventory and support status.", "Define core/edge/AP target architecture."],
        d90: ["Start high-priority edge and AP replacement waves."],
        y12: ["Deliver campus-wide Wi-Fi 7 aligned network posture."],
      },
    },
    7: {
      currentState: [
        `${client.windowsDevices} Windows devices with ${client.oldOsTotal} legacy OS cohort.`,
        `${client.chromebooks} Chromebooks with ${client.chromeExpired} expired lifecycle entries.`,
        `${client.tabletDevices} tablets requiring policy/lifecycle governance consistency.`,
      ],
      whatThisMeans: "Client risk is driven by supportability and management maturity; unsupported OS devices should be prioritised for replacement.",
      priorityActions: {
        now: ["Identify unsupported OS devices and replacement order.", "Validate ownership and management policy baselines."],
        d90: ["Reduce high-risk endpoint cohorts by planned refresh."],
        y12: ["Standardise endpoint governance across all device classes."],
      },
    },
    8: {
      currentState: [
        `Cyber controls complete: ${cyberPct}.`,
        `${cyber.incompleteCount} controls incomplete; ${cyber.naCount} marked N/A.`,
        `Email security status: ${cyber.emailSecurityStatus}; cloud backup status: ${cyber.cloudBackupStatus}.`,
      ],
      whatThisMeans: "Control presence must be backed by evidence, testing and ownership to provide reliable resilience and assurance.",
      priorityActions: {
        now: ["Close incomplete/N/A controls with named owners.", "Validate backup and DR evidence with test outcomes."],
        d90: ["Standardise control review cadence and reporting."],
        y12: ["Align cyber baseline with trust-wide assurance framework."],
      },
    },
    9: {
      currentState: [
        `Current licence recorded as ${core.currentLicence}.`,
        `${core.configuredUnknown} configured capabilities have unclear usage evidence.`,
        `${core.featureRows} feature rows assessed in baseline review.`,
      ],
      whatThisMeans: "Value realisation depends on proving usage for enabled capability and enabling priority features that are currently under-used.",
      priorityActions: {
        now: ["Validate enabled features against actual usage.", "Prioritise high-value A3 capabilities for activation."],
        d90: ["Track adoption and governance outcomes by feature group."],
        y12: ["Decide licence uplift only after baseline optimisation evidence."],
      },
    },
    10: {
      currentState: [
        "Phased actions defined across immediate, 90-day and 12-month horizons.",
        `Migration remediation open count: ${migration.remediation}.`,
        `Cyber incomplete control count: ${cyber.incompleteCount}.`,
      ],
      whatThisMeans: "Execution success now depends on disciplined ownership, sequencing and evidence-driven governance rather than further discovery alone.",
      priorityActions: {
        now: ["Confirm owners and deadlines for all immediate actions."],
        d90: ["Deliver roadmap checkpoints and variance reporting."],
        y12: ["Close strategic gaps and embed annual lifecycle governance."],
      },
    },
  };

  function renderPdfExpansion(model, startY) {
    if (!model) return startY;
    let y = startY;
    const bottomLimit = pageHeight - 42;
    const ensureSpace = (required) => {
      if (y + required > bottomLimit) {
        doc.addPage();
        y = 44;
      }
    };
    const sectionHeading = (title) => {
      ensureSpace(18);
      doc.setTextColor(8, 35, 63);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.text(title, 40, y);
      y += 14;
    };
    const sectionParagraph = (text) => {
      const lines = doc.splitTextToSize(text, pageWidth - 88);
      ensureSpace((lines.length * 12) + 6);
      doc.setTextColor(36, 56, 79);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10.5);
      doc.text(lines, 48, y);
      y += (lines.length * 12) + 4;
    };
    const sectionBullets = (items, fallback) => {
      const list = (items || []).length ? items : [fallback];
      for (const item of list) {
        const bulletLines = doc.splitTextToSize(`• ${item}`, pageWidth - 96);
        ensureSpace((bulletLines.length * 12) + 2);
        doc.setTextColor(36, 56, 79);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10.5);
        doc.text(bulletLines, 48, y);
        y += (bulletLines.length * 12) + 2;
      }
    };

    sectionHeading("Current State");
    sectionBullets(model.currentState, "Current-state evidence is still being finalised.");
    y += 8;
    sectionHeading("What This Means");
    sectionParagraph(model.whatThisMeans || "This area needs structured remediation and clearer ownership before it can be considered low risk.");
    y += 8;
    sectionHeading("Priority Actions");

    const actions = model.priorityActions || {};
    const actionRows = [
      ["Now", (actions.now && actions.now.length) ? actions.now.join("\n• ") : "Define actions and owners for this period."],
      ["90 days", (actions.d90 && actions.d90.length) ? actions.d90.join("\n• ") : "Define actions and owners for this period."],
      ["12 months", (actions.y12 && actions.y12.length) ? actions.y12.join("\n• ") : "Define actions and owners for this period."],
    ];
    ensureSpace(42);
    doc.autoTable({
      head: [["Timeframe", "Priority actions"]],
      body: actionRows,
      startY: y,
      margin: { left: 28, right: 28 },
      styles: { font: "helvetica", fontSize: 9.5, cellPadding: 5, lineColor: [212, 219, 230], lineWidth: 0.6, textColor: [27, 31, 36], overflow: "linebreak" },
      headStyles: { fillColor: [22, 56, 93], textColor: [255, 255, 255], fontStyle: "bold" },
      alternateRowStyles: { fillColor: [238, 243, 249] },
      theme: "grid",
    });
    return (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 10 : y + 10;
  }

  for (let i = 0; i < chapters.length; i++) {
    const ch = chapters[i];
    doc.addPage();
    const chapterTitleLines = doc.splitTextToSize(ch.title, pageWidth - 80);
    const chapterHeadHeight = Math.max(56, 26 + (chapterTitleLines.length * 22));
    doc.setFillColor(8, 35, 63);
    doc.rect(28, 24, pageWidth - 56, chapterHeadHeight, "F");
    doc.setTextColor(125, 211, 252);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text(`CHAPTER ${i + 1}`, 40, 43);
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(22);
    doc.text(chapterTitleLines, 40, 66);

    let y = 24 + chapterHeadHeight + 18;
    if (ch.intro) {
      doc.setTextColor(27, 31, 36);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(11.5);
      const lines = doc.splitTextToSize(ch.intro, pageWidth - 80);
      doc.text(lines, 40, y);
      y += (lines.length * 14) + 8;
    }

    doc.autoTable({
      head: [ch.head],
      body: ch.body.length ? ch.body : [["-", "-", "-", "-"].slice(0, ch.head.length)],
      startY: y,
      margin: { left: 28, right: 28 },
      styles: { font: "helvetica", fontSize: 10, cellPadding: 6, lineColor: [212, 219, 230], lineWidth: 0.6, textColor: [27, 31, 36], overflow: "linebreak" },
      headStyles: { fillColor: [22, 56, 93], textColor: [255, 255, 255], fontStyle: "bold" },
      alternateRowStyles: { fillColor: [238, 243, 249] },
      theme: "grid",
    });

    let followOnY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 14 : y + 14;

    if (i === 0) {
      if (followOnY > pageHeight - 120) {
        doc.addPage();
        followOnY = 44;
      }
      doc.setTextColor(8, 35, 63);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.text("Unleashed Brilliant Basics", 40, followOnY);
      const basicsBody = brilliantBasicsRows.map((row) => [row.capability, row.status, row.evidence]);
      doc.autoTable({
        head: [["Control", "Status", "Evidence note"]],
        body: basicsBody,
        startY: followOnY + 8,
        margin: { left: 28, right: 28 },
        styles: { font: "helvetica", fontSize: 9.5, cellPadding: 5, lineColor: [212, 219, 230], lineWidth: 0.6, textColor: [27, 31, 36], overflow: "linebreak" },
        headStyles: { fillColor: [22, 56, 93], textColor: [255, 255, 255], fontStyle: "bold" },
        alternateRowStyles: { fillColor: [238, 243, 249] },
        theme: "grid",
      });
      followOnY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 10 : followOnY + 10;
    }

    if (i >= 3) {
      const expansionModel = pdfExpansionByChapter[i + 1];
      followOnY = renderPdfExpansion(expansionModel, followOnY);
    }
  }

  doc.save("audit-dashboard-export.pdf");
}

async function ensurePdfLibrariesLoaded() {
  if (window.jspdf && window.jspdf.jsPDF && window.jspdf.jsPDF.API.autoTable) return;
  await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", "jspdf-lib");
  await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js", "jspdf-autotable-lib");
}

function loadScriptOnce(src, id) {
  return new Promise((resolve, reject) => {
    const existing = document.getElementById(id);
    if (existing) {
      if (existing.getAttribute("data-loaded") === "true") resolve();
      else existing.addEventListener("load", () => resolve(), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.id = id;
    script.src = src;
    script.async = true;
    script.onload = () => {
      script.setAttribute("data-loaded", "true");
      resolve();
    };
    script.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(script);
  });
}

function buildExportHtml(mode = "web") {
  const data = latestReport;
  if (!data) {
    return `<!doctype html><html><head><meta charset="utf-8"><title>Audit Export</title></head><body><p>No workbook loaded.</p></body></html>`;
  }
  const isWord = mode === "word";
  const isPdf = mode === "pdf";

  const { infra, network, client, cyber, migration, core, dashboard } = data;
  const software = data.software || { entries: [] };
  const lifecycle = mergeLifecycleStores([infra.lifecycle, network.lifecycle, client.lifecycle]);
  const endWarranty = (lifecycle.flagged.warranty || []).length;
  const endSupport = (lifecycle.flagged.support || []).length;
  const endServiceLife = (lifecycle.flagged.serviceLife || []).length;
  const msReadyPct = migration.total ? `${((migration.ready / migration.total) * 100).toFixed(1)}%` : "-";
  const cyberPctFallback = cyber.totalControls > 0 ? `${((cyber.completeCount / cyber.totalControls) * 100).toFixed(1)}%` : "-";
  const cyberPct = dashboard.cyberPct !== "-" ? dashboard.cyberPct : cyberPctFallback;
  const thankYouLine = "Thank you for undertaking the Unleashed site audit and readiness assessment.";

  const infraStatus = (infra.ws2012 > 0 || infra.serverCritical > 0) ? "Red" : "Amber";
  const networkStatus = (network.apsOutOfWarranty + network.switchesOutOfWarranty + network.firewallsOutOfWarranty) > 0 ? "Red" : "Amber";
  const clientStatus = client.oldOsTotal > 0 ? "Amber" : "Green";
  const cyberStatus = cyber.incompleteCount > 0 ? "Amber / Red" : "Green";
  const tenancyStatus = migration.remediation > 0 ? "Red" : "Amber";
  const coreStatus = core.configuredUnknown > 0 ? "Amber" : "Green";

  const remediationRows = (migration.items || [])
    .filter((item) => item.stage === "Remediation required")
    .slice(0, 8)
    .map((item) => `
      <tr>
        <td>${escapeHtml(item.area)}</td>
        <td>${escapeHtml(item.stage)}</td>
        <td>${escapeHtml(item.check)}</td>
        <td>${escapeHtml(item.action)}</td>
      </tr>
    `)
    .join("");

  const brilliantBasicsRows = buildBrilliantBasicsRows(cyber, core, software);

  const brilliantBasicsRowsHtml = brilliantBasicsRows
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.capability)}</td>
        <td>${escapeHtml(row.status)}</td>
        <td>${escapeHtml(row.evidence)}</td>
      </tr>
    `)
    .join("");

  const sharedTopAssets = [
    { asset: "Physical server estate", ageSupport: `${infra.totalServers} servers / ${infra.ws2012} legacy OS`, risk: infraStatus, path: "Retain only critical local workloads and migrate feasible services to SaaS/Azure." },
    { asset: "Network switching", ageSupport: `Core ${network.core}, Edge ${network.edge}`, risk: networkStatus, path: "Phase core-first and edge refresh to support Wi-Fi 7 and modern uplinks." },
    { asset: "Wireless platform", ageSupport: `${network.aps} APs with lifecycle pressure`, risk: networkStatus, path: "Prioritise high-density areas and align AP refresh with switching/PoE readiness." },
    { asset: "Client compute", ageSupport: `${client.totalQuantity || dashboard.devices} devices / ${client.oldOsTotal} legacy OS`, risk: clientStatus, path: "Replace unsupported OS cohorts first and standardise endpoint policy control." },
    { asset: "Cyber controls", ageSupport: `${cyberPct} complete / ${cyber.incompleteCount} incomplete`, risk: cyberStatus, path: "Close incomplete controls and evidence recovery/backup effectiveness." },
  ];

  function renderBulletList(items, fallback) {
    const list = (items || []).filter((item) => String(item || "").trim() !== "");
    if (!list.length) return `<ul><li>${escapeHtml(fallback)}</li></ul>`;
    return `<ul>${list.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }

  function renderPriorityActionsTable(priorityActions) {
    const actions = priorityActions || {};
    const byPeriod = [
      { label: "Now", items: actions.now || [] },
      { label: "90 days", items: actions.d90 || [] },
      { label: "12 months", items: actions.y12 || [] },
    ];
    return `
      <table class="mini-table">
        <thead><tr><th>Timeframe</th><th>Priority actions</th></tr></thead>
        <tbody>
          ${byPeriod.map((row) => `
            <tr>
              <td>${row.label}</td>
              <td>${renderBulletList(row.items, "Define actions and owners for this period.")}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function renderDependenciesOwnersTable(rows) {
    const deps = (rows || []).filter((row) => row && (row.dependency || row.owner || row.reason));
    const safe = deps.length ? deps : [{ dependency: "Dependency mapping pending", owner: "School IT + Trust Programme", reason: "Confirm operational and governance ownership." }];
    return `
      <table class="mini-table">
        <thead><tr><th>Dependency</th><th>Owner</th><th>Why this matters</th></tr></thead>
        <tbody>
          ${safe.map((row) => `
            <tr>
              <td>${escapeHtml(row.dependency || "-")}</td>
              <td>${escapeHtml(row.owner || "-")}</td>
              <td>${escapeHtml(row.reason || "-")}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function renderTopAssetsTable(rows) {
    const assets = (rows || []).slice(0, 5);
    const safe = assets.length ? assets : [{ asset: "No top assets captured", ageSupport: "To be validated", risk: "Medium", path: "Confirm lifecycle and replacement path." }];
    return `
      <table class="mini-table">
        <thead><tr><th>Asset / system</th><th>Age / support state</th><th>Risk</th><th>Replacement path</th></tr></thead>
        <tbody>
          ${safe.map((row) => `
            <tr>
              <td>${escapeHtml(row.asset || "-")}</td>
              <td>${escapeHtml(row.ageSupport || "-")}</td>
              <td>${escapeHtml(row.risk || "-")}</td>
              <td>${escapeHtml(row.path || "-")}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function renderBeforeTargetTable(rows) {
    const states = (rows || []);
    const safe = states.length ? states : [{ area: "Current baseline", before: "Evidence pending", target: "Validated and approved target-state definition." }];
    return `
      <table class="mini-table">
        <thead><tr><th>Area</th><th>Before</th><th>Target</th></tr></thead>
        <tbody>
          ${safe.map((row) => `
            <tr>
              <td>${escapeHtml(row.area || "-")}</td>
              <td>${escapeHtml(row.before || "-")}</td>
              <td>${escapeHtml(row.target || "-")}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function renderChapterExpansion(model) {
    const m = model || {};
    const sections = [
      {
        title: "Current State",
        content: renderBulletList(m.currentState, "Current-state evidence is still being finalised."),
      },
      {
        title: "What This Means",
        content: `<p class="narrative">${escapeHtml(m.whatThisMeans || "This area needs structured remediation and clearer ownership before it can be considered low risk.")}</p>`,
      },
      {
        title: "Priority Actions (Now / 90 days / 12 months)",
        content: renderPriorityActionsTable(m.priorityActions),
      },
    ];
    return `
      <div class="expansion-section">
        ${sections.map((section) => `
          <div class="expansion-block">
            <h3>${section.title}</h3>
            ${section.content}
          </div>
        `).join("")}
      </div>
    `;
  }

  const chapterModels = {
    executive: {
      currentState: [
        `${infra.totalServers} server records with lifecycle pressure in core infrastructure.`,
        `Cyber completion at ${cyberPct} with ${cyber.incompleteCount} controls requiring closure.`,
        `Microsoft migration readiness currently ${msReadyPct} with ${migration.remediation} remediation actions.`,
        `Client estate includes ${client.oldOsTotal} legacy OS devices requiring phased replacement.`,
      ],
      whatThisMeans: "Leadership can proceed with migration and refresh planning only where critical remediation is actively managed and evidenced.",
      priorityActions: { now: ["Confirm governance owners for each critical risk.", "Publish remediation tracker and executive checkpoints."], d90: ["Close migration blockers and validate DR evidence."], y12: ["Complete lifecycle-led modernisation roadmap."] },
      dependencies: [{ dependency: "Trust governance cadence", owner: "Trust IT + School SLT", reason: "Decision velocity and risk acceptance must be formalised." }],
      riskIfDelayed: "Critical milestones may be missed, forcing reactive replacements and increasing disruption risk.",
      costEffort: { band: "High", rationale: "Multiple workstreams span infrastructure, cyber, migration and end-user devices." },
      successCriteria: ["Monthly reduction in remediation backlog.", "Migration go/no-go based on evidence, not assumptions."],
      assumptions: ["Dashboard source metrics are current and validated.", "Named owners can commit delivery capacity."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Governance", before: "Reactive and assumption-led decision making.", target: "Evidence-led monthly governance with clear accountability." }, { area: "Migration readiness", before: `${msReadyPct} ready with open blockers.`, target: "Critical blockers closed and readiness sign-off completed." }],
    },
    rag: {
      currentState: ["Data, network and tenancy readiness remain red/amber risk areas.", `Cyber remains ${cyberStatus} with incomplete control coverage.`, "Core application optimisation is still under-used."],
      whatThisMeans: "The RAG profile indicates concentrated operational and migration risk that should drive prioritised investment and ownership.",
      priorityActions: { now: ["Treat red items as board-tracked actions.", "Agree risk acceptance thresholds for amber items."], d90: ["Move red items to amber through evidence-backed remediation."], y12: ["Stabilise all high-risk domains into managed lifecycle state."] },
      dependencies: [{ dependency: "Accurate tab-level inputs", owner: "Site IT + Audit owner", reason: "RAG quality depends on validated workbook data." }],
      riskIfDelayed: "RAG status drifts from reality and leadership decisions may understate true risk.",
      costEffort: { band: "Medium", rationale: "Primarily governance and remediation tracking effort with targeted technical changes." },
      successCriteria: ["Red domains have dated mitigation plans.", "Amber domains have owners and evidence checkpoints."],
      assumptions: ["RAG labels are interpreted consistently by stakeholders."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Risk visibility", before: "Static status snapshots.", target: "Live risk movement with owner-level accountability." }],
    },
    lifecycle: {
      currentState: [`Warranty flags: ${endWarranty}.`, `End-of-support flags (critical): ${endSupport}.`, `End-of-service-life flags: ${endServiceLife}.`],
      whatThisMeans: "Lifecycle planning must be linked to supportability risk, not just hardware age or break/fix history.",
      priorityActions: { now: ["Validate support contracts and true lifecycle dates.", "Prioritise end-of-support assets for remediation."], d90: ["Agree extension vs replacement decisions."], y12: ["Embed rolling lifecycle governance and budget sequencing."] },
      dependencies: [{ dependency: "Vendor lifecycle confirmation", owner: "IT Operations", reason: "Accurate support dates drive replacement priority." }],
      riskIfDelayed: "Unsupported assets may fail without viable support paths, causing unplanned cost and outage risk.",
      costEffort: { band: "Medium", rationale: "Assessment effort is moderate; replacement work can be high where scope is large." },
      successCriteria: ["All critical support-end assets have approved treatment path.", "Lifecycle register is maintained as live control."],
      assumptions: ["N/A lifecycle values are treated as risk until validated."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Lifecycle model", before: "Assumption-led expiry posture.", target: "Validated dates with approved extension/replacement decisions." }],
    },
    migration: {
      currentState: [`Microsoft readiness currently ${msReadyPct}.`, `${migration.remediation} remediation-required migration checks remain open.`, `${migration.inProgress} actions are in progress and need owner follow-through.`],
      whatThisMeans: "Cutover confidence depends on closing remediation items and validating dependencies across identity, mail, collaboration and apps.",
      priorityActions: { now: ["Close remediation-required rows and confirm evidence.", "Validate DNS, mail flow and identity controls."], d90: ["Complete UAT and go/no-go decision pack."], y12: ["Transition to post-migration optimisation and decommission governance."] },
      dependencies: [{ dependency: "Tenant admin exports and evidence pack", owner: "Migration lead", reason: "Readiness cannot be approved without complete evidence." }],
      riskIfDelayed: "Migration timeline may slip or proceed with unresolved blockers, increasing operational risk.",
      costEffort: { band: "High", rationale: "Cross-domain remediation and coordination with trust and school teams are required." },
      successCriteria: ["All red remediation rows closed or approved by exception.", "Go/no-go decision completed with full evidence pack."],
      assumptions: ["Source platform and dependency inventory remain accurate during delivery."],
      topAssets: (migration.items || []).slice(0, 5).map((item) => ({ asset: item.area, ageSupport: item.status, risk: item.stage, path: item.action })),
      beforeTarget: [{ area: "Readiness", before: `${msReadyPct} with unresolved remediation.`, target: "Critical actions closed and migration approval granted." }],
    },
    dataInfra: {
      currentState: [`${infra.totalServers} records with ${infra.ws2012} legacy workloads.`, `${infra.serverCritical} immediate server candidates identified.`, `${infra.sanSwitchCount} SAN switches and ${infra.sanStorageCount} storage arrays require validation.`],
      whatThisMeans: "Server and storage strategy should prioritise unsupported workloads and reduce local dependency where cloud alternatives are viable.",
      priorityActions: { now: ["Remediate unsupported server OS workloads.", "Confirm SAN/storage support state and dependency mapping."], d90: ["Agree retain vs migrate vs replace decisions."], y12: ["Execute phased refresh only for services that must remain local."] },
      dependencies: [{ dependency: "Application dependency map", owner: "Infrastructure lead", reason: "Determines which services can move off local compute." }],
      riskIfDelayed: "Critical services may remain on unsupported platforms with weaker recovery confidence.",
      costEffort: { band: "High", rationale: "Infrastructure remediation and workload migration can be capital and effort intensive." },
      successCriteria: ["Unsupported workloads have approved migration/remediation path.", "Storage and backup dependencies are validated."],
      assumptions: ["Host vs VM lifecycle interpretation remains consistent."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Server estate", before: "Mixed support state and legacy dependency.", target: "Supported core workloads with cloud-first reduction plan." }],
    },
    network: {
      currentState: [`Core ${network.core} and edge ${network.edge} switching in mixed lifecycle state.`, `${network.aps} APs with Wi-Fi 7 readiness gap.`, `${network.firewalls} firewall records requiring support-path validation.`],
      whatThisMeans: "Wi-Fi 7 readiness requires coordinated core, edge and AP modernization rather than isolated hardware swaps.",
      priorityActions: { now: ["Validate network inventory and support status.", "Define core/edge/AP target architecture."], d90: ["Start high-priority edge and AP replacement waves."], y12: ["Deliver campus-wide Wi-Fi 7 aligned network posture."] },
      dependencies: [{ dependency: "PoE and uplink design standards", owner: "Network lead", reason: "Switching design governs AP capability and performance." }],
      riskIfDelayed: "Wireless performance and security posture may degrade while replacement cost and complexity increase.",
      costEffort: { band: "High", rationale: "Campus network transformation involves significant hardware and implementation effort." },
      successCriteria: ["Core and edge standards approved (10Gb core / 2.5Gb edge).", "Wi-Fi refresh aligned with switching dependencies."],
      assumptions: ["AP and switch inventory reflects active estate and location accuracy."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Wireless capability", before: "Mixed Wi-Fi generation with lifecycle pressure.", target: "Planned Wi-Fi 7-capable architecture with aligned switching." }],
    },
    client: {
      currentState: [`${client.windowsDevices} Windows devices with ${client.oldOsTotal} legacy OS cohort.`, `${client.chromebooks} Chromebooks with ${client.chromeExpired} expired lifecycle entries.`, `${client.tabletDevices} tablets requiring policy/lifecycle governance consistency.`],
      whatThisMeans: "Client risk is driven by supportability and management maturity; unsupported OS devices should be prioritised for replacement.",
      priorityActions: { now: ["Identify unsupported OS devices and replacement order.", "Validate ownership and management policy baselines."], d90: ["Reduce high-risk endpoint cohorts by planned refresh."], y12: ["Standardise endpoint governance across all device classes."] },
      dependencies: [{ dependency: "Asset register accuracy", owner: "Endpoint lead", reason: "Replacement and policy decisions depend on trusted inventory." }],
      riskIfDelayed: "Unsupported devices increase security and support overhead while reducing operational resilience.",
      costEffort: { band: "Medium", rationale: "Effort is sustained and operationally heavy, with phased hardware spend." },
      successCriteria: ["Legacy OS cohort reduced quarter-on-quarter.", "Policy and compliance baseline applied consistently."],
      assumptions: ["Device ownership and usage context are correctly recorded."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Endpoint posture", before: "Mixed supportability and management maturity.", target: "Supported OS baseline with consistent policy enforcement." }],
    },
    cyber: {
      currentState: [`Cyber controls complete: ${cyberPct}.`, `${cyber.incompleteCount} controls incomplete; ${cyber.naCount} marked N/A.`, `Email security status: ${cyber.emailSecurityStatus}; cloud backup status: ${cyber.cloudBackupStatus}.`],
      whatThisMeans: "Control presence must be backed by evidence, testing and ownership to provide reliable resilience and assurance.",
      priorityActions: { now: ["Close incomplete/N/A controls with named owners.", "Validate backup and DR evidence with test outcomes."], d90: ["Standardise control review cadence and reporting."], y12: ["Align cyber baseline with trust-wide assurance framework."] },
      dependencies: [{ dependency: "Evidence collection discipline", owner: "Security lead", reason: "Assurance quality depends on repeatable evidence capture." }],
      riskIfDelayed: "Incident response confidence remains low and recovery outcomes may not meet expectations.",
      costEffort: { band: "Medium", rationale: "Mostly process and control hardening effort with targeted tooling adjustments." },
      successCriteria: ["All critical controls evidenced and tested.", "Backup/DR controls verified against defined RTO/RPO."],
      assumptions: ["Control status labels align with real operational state."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Assurance posture", before: "Tool-centric with partial evidence.", target: "Evidence-backed controls with tested recovery outcomes." }],
    },
    core: {
      currentState: [`Current licence recorded as ${core.currentLicence}.`, `${core.configuredUnknown} configured capabilities have unclear usage evidence.`, `${core.featureRows} feature rows assessed in baseline review.`],
      whatThisMeans: "Value realisation depends on proving usage for enabled capability and enabling priority features that are currently under-used.",
      priorityActions: { now: ["Validate enabled features against actual usage.", "Prioritise high-value A3 capabilities for activation."], d90: ["Track adoption and governance outcomes by feature group."], y12: ["Decide licence uplift only after baseline optimisation evidence." ] },
      dependencies: [{ dependency: "Tenant configuration evidence", owner: "M365 platform owner", reason: "Configuration and usage proof drives optimisation decisions." }],
      riskIfDelayed: "Licence value remains under-realised and overlapping tooling costs persist.",
      costEffort: { band: "Low", rationale: "Primarily configuration validation, governance and adoption enablement." },
      successCriteria: ["Configured features have owner + usage evidence.", "Priority A3 opportunities moved to operational use."],
      assumptions: ["Current licence and feature matrix remain accurate."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Licence value", before: "Enabled state not consistently evidenced.", target: "Usage-proven features and clear optimisation roadmap." }],
    },
    final: {
      currentState: ["Phased actions defined across immediate, 90-day and 12-month horizons.", `Migration remediation open count: ${migration.remediation}.`, `Cyber incomplete control count: ${cyber.incompleteCount}.`],
      whatThisMeans: "Execution success now depends on disciplined ownership, sequencing and evidence-driven governance rather than further discovery alone.",
      priorityActions: { now: ["Confirm owners and deadlines for all immediate actions."], d90: ["Deliver roadmap checkpoints and variance reporting."], y12: ["Close strategic gaps and embed annual lifecycle governance."] },
      dependencies: [{ dependency: "Programme governance and reporting cadence", owner: "Trust + School leadership", reason: "Sustained progress requires active oversight and decision-making." }],
      riskIfDelayed: "Roadmap momentum reduces and unresolved risks carry forward into future audit cycles.",
      costEffort: { band: "Medium", rationale: "Execution is primarily coordination-heavy with targeted technical delivery." },
      successCriteria: ["Roadmap milestones achieved on planned cadence.", "Risk exposure reduced in each review cycle."],
      assumptions: ["Resourcing remains available for planned remediation timeline."],
      topAssets: sharedTopAssets,
      beforeTarget: [{ area: "Delivery posture", before: "Action list without sustained governance rhythm.", target: "Managed programme with measurable milestone completion." }],
    },
  };

  const chapterPage = (n, title, content, isFirst = false) => `
    ${isWord && !isFirst ? `<br class="word-page-break" clear="all" style="mso-special-character:line-break;page-break-before:always;">` : ""}
    ${isWord && !isFirst ? `<div class="chapter-spacer"><span>Chapter break</span></div>` : ""}
    <section class="chapter-page${isFirst ? " chapter-first" : ""}">
      <div class="chapter-head">
        <div class="kicker">CHAPTER ${n}</div>
        <h2>${escapeHtml(title)}</h2>
      </div>
      ${content}
    </section>
  `;

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Audit Dashboard Export</title>
  <style>
    @page { size: A4; margin: 18mm 14mm 18mm 14mm; }
    body { font-family: Calibri, Arial, sans-serif; margin: 0; color: #1b1f24; background: ${isWord || isPdf ? "#ffffff" : "#eef2f7"}; }
    .page { margin: 0 auto; max-width: 1020px; padding: 18px 14px 32px; }
    .cover {
      background: ${isWord ? "#0a3058" : "linear-gradient(150deg, #07213d, #0a3058 65%, #195085)"};
      color: #fff;
      border-radius: 8px;
      padding: 36px 34px 30px;
      min-height: 380px;
      break-after: page;
      page-break-after: always;
      mso-page-break-after: always;
    }
    .brand { font-size: 42px; font-weight: 700; letter-spacing: 1px; margin: 0; }
    .brand-sub { margin-top: 4px; color: #dbeafe; font-size: 15px; }
    .cover-title { margin-top: 54px; font-size: ${isWord ? "38px" : "44px"}; line-height: 1.08; font-weight: 700; overflow-wrap: anywhere; }
    .cover-school { margin-top: 8px; color: #7dd3fc; font-size: 30px; font-weight: 700; }
    .cover-copy { margin-top: 20px; font-size: 18px; max-width: 820px; color: #dbeafe; }
    .cover-focus { margin-top: 26px; font-size: 20px; font-weight: 700; }
    .cover-focus-sub { margin-top: 6px; color: #93c5fd; }

    .meta-line { margin: 12px 2px 16px; color: #5a6f8b; font-size: 13px; }
    .chapter-page {
      margin-top: 18px;
      break-before: auto;
      page-break-before: auto;
      mso-page-break-before: auto;
      break-inside: avoid;
      page-break-inside: avoid;
      mso-page-break-inside: avoid;
    }
    .word-export .chapter-page {
      break-before: page;
      page-break-before: always;
      mso-page-break-before: always;
      break-inside: auto;
      page-break-inside: auto;
      mso-page-break-inside: auto;
    }
    .word-export .chapter-page.chapter-first {
      break-before: auto;
      page-break-before: auto;
      mso-page-break-before: auto;
    }
    .word-page-break {
      display: block;
      line-height: 0;
      font-size: 0;
      page-break-before: always;
      break-before: page;
      mso-page-break-before: always;
      mso-special-character: line-break;
      height: 0;
      margin: 0;
      border: 0;
    }
    .chapter-spacer { display: none; }
    .word-export .chapter-spacer {
      display: block;
      height: 48px;
      margin: 18px 0 10px;
      border-top: 2px dashed #9db1cb;
      position: relative;
    }
    .word-export .chapter-spacer span {
      position: absolute;
      top: 8px;
      right: 0;
      font-size: 11px;
      color: #5a6f8b;
      letter-spacing: 0.4px;
      text-transform: uppercase;
    }
    @media print {
      .cover { break-after: page !important; page-break-after: always !important; }
      .word-page-break { break-before: page !important; page-break-before: always !important; }
      body { background: #ffffff !important; }
    }
    .chapter-head { background: #08233F; color: #fff; padding: 10px 12px; border-radius: 4px; }
    .kicker { font-size: 13px; color: #7dd3fc; letter-spacing: 0.5px; margin-bottom: 3px; }
    .chapter-head h2 { margin: 0; font-size: ${isWord ? "28px" : "34px"}; line-height: 1.1; overflow-wrap: anywhere; }
    h3 { margin: 14px 0 8px; color: #08233F; font-size: 22px; }
    p { margin: 8px 0; line-height: 1.42; }
    .narrative { color: #24384f; }
    .expansion-block { margin-bottom: 14px; }
    .expansion-block:last-child { margin-bottom: 0; }

    table { width: 100%; border-collapse: collapse; margin-top: 8px; table-layout: ${isWord ? "auto" : "fixed"}; ${isWord ? "" : "box-shadow: 0 2px 8px rgba(8,35,63,0.08);"} }
    th, td { border: 1px solid #d4dbe6; padding: 8px; text-align: left; vertical-align: top; }
    th { background: #16385d; color: #fff; font-weight: 700; }
    tr:nth-child(even) td { background: #eef3f9; }
    thead { display: table-header-group; }
    tr, td, th { break-inside: avoid; page-break-inside: avoid; mso-page-break-inside: avoid; }
    ul { margin-top: 8px; }
    .section-gap { margin-top: 14px; }
  </style>
</head>
<body class="${isWord ? "word-export" : (isPdf ? "pdf-export" : "web-export")}">
  <div class="page">
    <section class="cover">
      <h1 class="brand">UNLEASHED</h1>
      <div class="brand-sub">Cyber Security & Managed IT</div>
      <div class="cover-title">IT Audit Summary, Findings and Roadmap</div>
      <div class="cover-school">Sir James Smith's School</div>
      <div class="cover-copy">${escapeHtml(thankYouLine)} This report provides a structured narrative of the current estate, key risks, decisions required and the recommended remediation roadmap.</div>
      <div class="cover-focus">Audit focus</div>
      <div class="cover-focus-sub">Infrastructure • Cyber Security • Migration Readiness</div>
    </section>

    <div class="meta-line">${escapeHtml(els.workbookMeta.textContent || "")} | Generated: ${new Date().toLocaleString()}</div>

    ${chapterPage(1, "Executive Summary", `
      <p class="narrative">This report translates the audit workbook into a board-level narrative with decisions, risks and a practical improvement roadmap.</p>
      <p class="narrative">It should be used as an executive decision document: what requires immediate action, what can be phased, and where evidence must be strengthened before major change.</p>
      <table>
        <thead><tr><th>Headline area</th><th>Position</th><th>Summary</th></tr></thead>
        <tbody>
          <tr><td>Infrastructure lifecycle</td><td>${infraStatus}</td><td>${infra.totalServers} servers recorded. Support-end and lifecycle flags require managed refresh planning.</td></tr>
          <tr><td>Cyber assurance</td><td>${cyberStatus}</td><td>Cyber completion ${cyberPct}; not complete controls ${cyber.incompleteCount}.</td></tr>
          <tr><td>Migration readiness (Microsoft)</td><td>${tenancyStatus}</td><td>${msReadyPct} ready, ${migration.remediation} remediation items remain.</td></tr>
          <tr><td>Core application / A3</td><td>${coreStatus}</td><td>Configured-but-usage-unknown capabilities: ${core.configuredUnknown}.</td></tr>
        </tbody>
      </table>
      <h3>Unleashed Brilliant Basics</h3>
      <table>
        <thead><tr><th>Control</th><th>Status</th><th>Evidence note</th></tr></thead>
        <tbody>${brilliantBasicsRowsHtml}</tbody>
      </table>
    `, true)}

    ${chapterPage(2, "RAG Summary", `
      <p class="narrative">The RAG summary highlights where risk is concentrated across infrastructure, network, cyber and migration readiness so governance can focus on the highest-impact gaps first.</p>
      <table>
        <thead><tr><th>Area</th><th>Status</th><th>Comment</th></tr></thead>
        <tbody>
          <tr><td>Data Infrastructure</td><td>${infraStatus}</td><td>Server lifecycle and OS risks require phased action.</td></tr>
          <tr><td>Network and Wi-Fi</td><td>${networkStatus}</td><td>Warranty and lifecycle flags across switching, APs and firewall estate.</td></tr>
          <tr><td>Client Compute</td><td>${clientStatus}</td><td>Unsupported/ageing OS cohorts require replacement planning.</td></tr>
          <tr><td>Cyber Security</td><td>${cyberStatus}</td><td>Key controls still not complete; cloud backup and assurance remain priorities.</td></tr>
          <tr><td>Tenancy Readiness</td><td>${tenancyStatus}</td><td>Remediation closure needed before migration go/no-go.</td></tr>
          <tr><td>Core Application / A3</td><td>${coreStatus}</td><td>A3 optimisation and feature utilisation review remains open.</td></tr>
        </tbody>
      </table>
    `)}

    ${chapterPage(3, "Lifecycle and Critical Milestones", `
      <p class="narrative">Lifecycle interpretation follows the trust principles: review at five years, replace by seven years, and apply cloud-first replacement wherever practical.</p>
      <p class="narrative">End of support is treated as the key critical milestone because it materially affects supportability, risk and resilience.</p>
      <table>
        <thead><tr><th>Milestone</th><th>Flagged count</th><th>Planning interpretation</th></tr></thead>
        <tbody>
          <tr><td>End of warranty</td><td>${endWarranty}</td><td>Can be extension candidates where support/service life still valid.</td></tr>
          <tr><td>End of support (critical)</td><td>${endSupport}</td><td>Critical replacement/upgrade trigger.</td></tr>
          <tr><td>End of service life</td><td>${endServiceLife}</td><td>Higher risk of reactive failure and diminishing support options.</td></tr>
        </tbody>
      </table>
    `)}

    ${chapterPage(4, "Migration Readiness – Actions", `
      <p class="narrative">This chapter intentionally shows remediation-only migration items so project focus remains on what must be closed before a safe and controlled cutover.</p>
      <table>
        <thead><tr><th>Area</th><th>Status</th><th>Main risk / gap</th><th>Immediate action</th></tr></thead>
        <tbody>${remediationRows || `<tr><td colspan="4">No open migration actions.</td></tr>`}</tbody>
      </table>
      ${renderChapterExpansion(chapterModels.migration)}
    `)}

    ${chapterPage(5, "Data Infrastructure – Servers, SAN and Storage", `
      <p class="narrative">Data infrastructure decisions should balance supportability, workload dependency and migration direction. The default approach should be to reduce local hosting where services can move to SaaS or Azure.</p>
      <table>
        <thead><tr><th>Infrastructure item</th><th>Position</th><th>Main concern</th><th>Recommended response</th></tr></thead>
        <tbody>
          <tr><td>Physical server estate</td><td>${infra.totalServers} records</td><td>${infra.ws2012} unsupported OS workloads and ${infra.serverCritical} immediate candidates.</td><td>Prioritise unsupported OS and support-end cohorts; phase cloud-first transition.</td></tr>
          <tr><td>SAN / storage</td><td>${infra.sanSwitchCount} SAN switches, ${infra.sanStorageCount} arrays</td><td>Operational dependency and lifecycle validation required.</td><td>Validate support, utilisation and target-state alignment.</td></tr>
          <tr><td>Virtual machines</td><td>${infra.virtual} virtual workloads</td><td>Guest OS risk is key (host lifecycle inherited from physical).</td><td>Treat unsupported VM OS as urgent remediation.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.dataInfra)}
    `)}

    ${chapterPage(6, "Network and Wi-Fi – Refresh Direction", `
      <p class="narrative">The network and wireless estate should be modernised as a single programme to achieve Wi-Fi 7 readiness, with core, edge and AP upgrades aligned by dependency.</p>
      <table>
        <thead><tr><th>Network area</th><th>Position</th><th>Main concern</th><th>Recommended response</th></tr></thead>
        <tbody>
          <tr><td>Core/edge switching</td><td>Core ${network.core}, Edge ${network.edge}</td><td>Lifecycle and warranty pressure across switching estate.</td><td>Phase refresh with core-first then edge modernisation.</td></tr>
          <tr><td>Wi-Fi platform</td><td>${network.aps} APs</td><td>Mix of Wi-Fi 5/6 and warranty exposure.</td><td>Design Wi-Fi 7 journey with switching/PoE alignment.</td></tr>
          <tr><td>Firewalls</td><td>${network.firewalls}</td><td>Supportability and replacement timing must be controlled.</td><td>Maintain supported edge security path.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.network)}
    `)}

    ${chapterPage(7, "Client Compute – Estate and Lifecycle", `
      <p class="narrative">For client compute, unsupported operating systems are the primary replacement trigger. Lifecycle governance should be tied to OS supportability rather than break/fix replacement alone.</p>
      <table>
        <thead><tr><th>Device area</th><th>Position</th><th>Main concern</th><th>Recommended response</th></tr></thead>
        <tbody>
          <tr><td>Windows</td><td>${client.windowsDevices}</td><td>Windows 10 / legacy cohorts: ${client.oldOsTotal} devices.</td><td>Plan staged replacement and supported OS compliance.</td></tr>
          <tr><td>Chromebooks</td><td>${client.chromebooks}</td><td>Expired ChromeOS cohorts: ${client.chromeExpired}.</td><td>Validate AUE lifecycle and replacement priority.</td></tr>
          <tr><td>Tablet / Mac / other</td><td>${client.tabletDevices + client.macDevices + client.otherDevices}</td><td>Consistency and governance quality vary by device class.</td><td>Improve ownership, OS visibility and policy standards.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.client)}
    `)}

    ${chapterPage(8, "Cyber Security Findings", `
      <p class="narrative">Cyber posture should be measured by evidenced control effectiveness, not only tooling presence. Controls marked N/A or incomplete should be prioritised for remediation and ownership.</p>
      <table>
        <thead><tr><th>Cyber area</th><th>Position</th><th>Main concern</th><th>Recommended response</th></tr></thead>
        <tbody>
          <tr><td>Cyber posture</td><td>${cyberPct} complete</td><td>${cyber.incompleteCount} controls not complete; ${cyber.naCount} marked N/A.</td><td>Close N/A/incomplete controls with clear ownership and evidence.</td></tr>
          <tr><td>Microsoft 365 backup</td><td>${escapeHtml(cyber.m365BackupStatus)}</td><td>Interim controls may be required pre-tenancy migration.</td><td>Confirm control ownership and delivery timeline.</td></tr>
          <tr><td>Email security / Cloud backup</td><td>${escapeHtml(cyber.emailSecurityStatus)} / ${escapeHtml(cyber.cloudBackupStatus)}</td><td>Cloud backup remains a critical resilience control.</td><td>Prioritise cloud backup and email security assurance.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.cyber)}
    `)}

    ${chapterPage(9, "Core Application and A3 Optimisation", `
      <p class="narrative">Core application and licensing review should focus on value realisation: configured capabilities must be validated as actively used, while non-enabled A3 features represent an avoidable ROI gap.</p>
      <table>
        <thead><tr><th>Area</th><th>Position</th><th>Main concern</th><th>Recommended response</th></tr></thead>
        <tbody>
          <tr><td>Current licence</td><td>${escapeHtml(core.currentLicence)}</td><td>Feature enablement may exceed proven operational usage.</td><td>Run usage validation and optimise A3 baseline.</td></tr>
          <tr><td>Configured-but-usage-unknown</td><td>${core.configuredUnknown}</td><td>Potential missed value and governance gap.</td><td>Map enabled features to active operational use.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.core)}
    `)}

    ${chapterPage(10, "Final Recommendations by Workbook Tab", `
      <p class="narrative">The phased plan below converts findings into an executable programme, sequencing urgent remediation first and strategic modernisation over the following 12 months.</p>
      <table>
        <thead><tr><th>Period</th><th>Area</th><th>Specific actions</th><th>Timing driver</th></tr></thead>
        <tbody>
          <tr><td>Immediate</td><td>Migration readiness</td><td>Close remediation actions and complete evidence pack for go/no-go.</td><td>${migration.remediation} remediation items open.</td></tr>
          <tr><td>0-3 months</td><td>Cyber security</td><td>Prioritise incomplete/N/A controls; evidence 365 backup, email security, cloud backup.</td><td>${cyber.incompleteCount} not complete controls.</td></tr>
          <tr><td>0-6 months</td><td>Data & network lifecycle</td><td>Address support-end/service-end flags and validate extension candidates.</td><td>${endSupport + endServiceLife} support/service-life flags.</td></tr>
          <tr><td>6-12 months</td><td>Client compute</td><td>Reduce unsupported OS cohorts and lifecycle exception devices.</td><td>${client.oldOsTotal} old OS devices.</td></tr>
        </tbody>
      </table>
      ${renderChapterExpansion(chapterModels.final)}
    `)}
  </div>
</body>
</html>`;
}
