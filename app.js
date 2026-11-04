/**
 * Ideathon 2025 - Problem Statement Selection Portal (VVCE)
 * Real-time frontend synchronization and concurrency handling
 */

// Application State
let problemStatements = [];
let currentDomainFilter = 'ALL';
let appsScriptUrl = localStorage.getItem('ivc_apps_script_url') || '';
let userClaimedPS = null;
let pollInterval = null;
let lastSyncTimestamp = Date.now();
let secondsSinceSync = 0;
let syncTimerInterval = null;
let pendingClaimPS = null;

// Initialize on DOM Load
document.addEventListener('DOMContentLoaded', () => {
  // Load saved claim if any
  try {
    const saved = localStorage.getItem('ivc_user_claimed_ps');
    if (saved) {
      userClaimedPS = JSON.parse(saved);
    }
  } catch (e) {
    console.error("Error reading saved claim", e);
  }

  // Load team info from localStorage if previously entered
  loadSavedTeamInfo();

  // Initialize initial dataset from data.js
  if (typeof INITIAL_PROBLEM_STATEMENTS !== 'undefined') {
    problemStatements = JSON.parse(JSON.stringify(INITIAL_PROBLEM_STATEMENTS));
  }

  // Set config URL input if saved
  const configInput = document.getElementById('configApiUrl');
  if (configInput && appsScriptUrl) {
    configInput.value = appsScriptUrl;
  }

  // Render initial UI
  updateClaimedBanner();
  renderProblemStatements();
  updateCounts();
  populateCustomDropdown();

  // Initialize Lucide Icons
  if (window.lucide) {
    lucide.createIcons();
  }

  // Start Real-Time Sync Engine
  startLiveSync();

  // Setup second counter for live badge
  syncTimerInterval = setInterval(() => {
    secondsSinceSync = Math.floor((Date.now() - lastSyncTimestamp) / 1000);
    const timerEl = document.getElementById('lastSyncTimer');
    if (timerEl) {
      timerEl.textContent = `${secondsSinceSync}s ago`;
    }
  }, 1000);

  // Close custom dropdown on outside click
  document.addEventListener('click', (e) => {
    const dropdown = document.getElementById('customDropdownMenu');
    const trigger = document.getElementById('customDropdownTrigger');
    if (dropdown && trigger && !trigger.contains(e.target) && !dropdown.contains(e.target)) {
      dropdown.classList.add('hidden');
      const arrow = document.getElementById('customDropdownArrow');
      if (arrow) arrow.classList.remove('rotate-180');
    }
  });
});

/**
 * START REAL-TIME LIVE SYNC
 * Polls Google Apps Script every 4 seconds for live concurrency updates
 */
function startLiveSync() {
  if (pollInterval) clearInterval(pollInterval);
  
  // First immediate sync
  refreshData(false);

  // Poll every 4 seconds
  pollInterval = setInterval(() => {
    refreshData(false);
  }, 4000);
}

/**
 * Fetch latest problem statements and availability status
 */
async function refreshData(manual = false) {
  const syncDot = document.getElementById('syncDot');
  const syncText = document.getElementById('syncStatusText');
  const refreshIcon = document.getElementById('refreshIcon');

  if (manual && refreshIcon) {
    refreshIcon.classList.add('animate-spin');
  }

  if (!appsScriptUrl) {
    // Demo Mode (Local Storage / Embedded data mode)
    if (syncDot) {
      syncDot.className = 'w-2.5 h-2.5 rounded-full bg-amber-400';
    }
    if (syncText) {
      syncText.textContent = 'Demo Mode (Local)';
    }
    
    // Check if any mock claims exist in localStorage
    syncLocalMockClaims();
    lastSyncTimestamp = Date.now();
    updateCounts();
    if (manual) {
      setTimeout(() => {
        if (refreshIcon) refreshIcon.classList.remove('animate-spin');
        showToast("Refreshed (Demo Mode). Add Google Sheet URL to enable live sync.", "info");
      }, 500);
    }
    return;
  }

  try {
    const url = appsScriptUrl.includes('?') 
      ? `${appsScriptUrl}&action=getPS&_t=${Date.now()}` 
      : `${appsScriptUrl}?action=getPS&_t=${Date.now()}`;

    const res = await fetch(url, { method: 'GET', mode: 'cors' });
    const json = await res.json();

    if (json && json.success && Array.isArray(json.data)) {
      // Merge live server status with our statement array
      const serverMap = new Map();
      json.data.forEach(item => {
        serverMap.set(item.code.toUpperCase(), item);
      });

      let changed = false;
      problemStatements.forEach(ps => {
        const live = serverMap.get(ps.code.toUpperCase());
        if (live) {
          const prevStatus = ps.status;
          const prevClaimedBy = ps.claimedBy;
          
          ps.status = live.status || (live.claimedBy ? 'Claimed' : 'Available');
          ps.claimedBy = live.claimedBy || '';
          ps.department = live.department || '';
          ps.claimedAt = live.claimedAt || '';

          if (prevStatus !== ps.status || prevClaimedBy !== ps.claimedBy) {
            changed = true;
          }
        }
      });

      lastSyncTimestamp = Date.now();

      if (syncDot) {
        syncDot.className = 'w-2.5 h-2.5 rounded-full bg-emerald-400 pulse-dot';
      }
      if (syncText) {
        syncText.textContent = 'Live Syncing';
      }

      // Re-render only if there are visual changes
      if (changed || manual) {
        renderProblemStatements();
        populateCustomDropdown();
      }
      updateCounts();

      // Check if user's claimed PS is verified on server
      if (userClaimedPS) {
        const myPs = problemStatements.find(p => p.code === userClaimedPS.psCode);
        if (myPs && myPs.claimedBy && myPs.claimedBy.toLowerCase() !== userClaimedPS.teamName.toLowerCase()) {
          // Warning: conflict detected
          showToast(`⚠️ Conflict: ${userClaimedPS.psCode} is held by ${myPs.claimedBy}`, "error");
        }
      }

      if (manual) {
        showToast("Synchronized with Google Sheets!", "success");
      }
    } else {
      throw new Error(json.message || "Invalid response format");
    }
  } catch (err) {
    console.warn("Sync error:", err);
    if (syncDot) {
      syncDot.className = 'w-2.5 h-2.5 rounded-full bg-rose-500';
    }
    if (syncText) {
      syncText.textContent = 'Sync Paused';
    }
    if (manual) {
      showToast("Sync error: Unable to reach Google Sheets Web App.", "error");
    }
  } finally {
    if (manual && refreshIcon) {
      setTimeout(() => refreshIcon.classList.remove('animate-spin'), 600);
    }
  }
}

/**
 * In demo mode, sync claims saved in localStorage so multiple tabs on the same computer reflect claims
 */
function syncLocalMockClaims() {
  try {
    const claims = JSON.parse(localStorage.getItem('ivc_mock_claims') || '{}');
    let changed = false;
    problemStatements.forEach(ps => {
      if (claims[ps.code]) {
        if (ps.status !== 'Claimed') changed = true;
        ps.status = 'Claimed';
        ps.claimedBy = claims[ps.code].teamName;
        ps.claimedAt = claims[ps.code].timestamp;
      }
    });
    if (changed) {
      renderProblemStatements();
      populateCustomDropdown();
    }
  } catch (e) {}
}

/**
 * Filter Problem Statements by Domain Tab
 */
function setDomainFilter(domain) {
  currentDomainFilter = domain;
  
  // Update Tab Styling
  const tabs = document.querySelectorAll('.domain-tab');
  tabs.forEach(tab => {
    if (tab.getAttribute('data-filter') === domain) {
      tab.className = 'domain-tab active px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 bg-blue-600 text-white shadow-md shadow-blue-500/20';
    } else {
      tab.className = 'domain-tab px-3.5 py-2 rounded-xl text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 transition flex items-center gap-1.5';
    }
  });

  renderProblemStatements();
}

/**
 * Render the Problem Statements Grid
 */
function renderProblemStatements() {
  const container = document.getElementById('psGrid');
  const emptyState = document.getElementById('emptyState');
  if (!container) return;

  const categoryFilter = document.getElementById('categorySelect').value;
  const statusFilter = document.getElementById('statusSelect').value;
  const searchQuery = (document.getElementById('searchInput').value || '').trim().toLowerCase();

  // Filter list
  const filtered = problemStatements.filter(ps => {
    // 1. Domain
    if (currentDomainFilter !== 'ALL' && ps.domain !== currentDomainFilter) {
      return false;
    }

    // 2. Category (Software / Hardware)
    if (categoryFilter !== 'ALL' && ps.category !== categoryFilter) {
      return false;
    }

    // 3. Status (Available / Claimed)
    const isClaimed = ps.status === 'Claimed' || !!ps.claimedBy;
    if (statusFilter === 'Available' && isClaimed) return false;
    if (statusFilter === 'Claimed' && !isClaimed) return false;

    // 4. Search query
    if (searchQuery) {
      const matchCode = ps.code.toLowerCase().includes(searchQuery);
      const matchSih = ps.sihNo.toLowerCase().includes(searchQuery);
      const matchTitle = ps.title.toLowerCase().includes(searchQuery);
      const matchOrg = (ps.org || '').toLowerCase().includes(searchQuery);
      const matchTheme = (ps.theme || '').toLowerCase().includes(searchQuery);
      const matchDesc = (ps.description || '').toLowerCase().includes(searchQuery);
      if (!matchCode && !matchSih && !matchTitle && !matchOrg && !matchTheme && !matchDesc) {
        return false;
      }
    }

    return true;
  });

  if (filtered.length === 0) {
    container.innerHTML = '';
    if (emptyState) emptyState.classList.remove('hidden');
    return;
  }

  if (emptyState) emptyState.classList.add('hidden');

  // Generate Cards HTML
  container.innerHTML = filtered.map(ps => {
    const isClaimed = ps.status === 'Claimed' || !!ps.claimedBy;
    const isMyClaim = userClaimedPS && userClaimedPS.psCode === ps.code;

    // Domain color styling
    let domainBadgeStyle = "bg-blue-500/10 text-blue-400 border-blue-500/20";
    if (ps.domain.includes("Electronics")) {
      domainBadgeStyle = "bg-purple-500/10 text-purple-400 border-purple-500/20";
    } else if (ps.domain.includes("Mechanical") || ps.domain.includes("Civil")) {
      domainBadgeStyle = "bg-emerald-500/10 text-emerald-400 border-emerald-500/20";
    }

    // Category badge styling
    const categoryBadgeStyle = ps.category === "Software" 
      ? "bg-indigo-500/20 text-indigo-300 border-indigo-500/30" 
      : "bg-amber-500/20 text-amber-300 border-amber-500/30";

    const categoryIcon = ps.category === "Software" ? "laptop" : "cpu";

    // Card boundary styling based on status
    let cardBorder = isClaimed 
      ? "border-slate-800 bg-slate-900/40 opacity-80" 
      : "border-slate-800 hover:border-slate-700 bg-slate-900/90 hover:shadow-xl hover:shadow-blue-500/5 transition duration-200";

    if (isMyClaim) {
      cardBorder = "border-emerald-500/60 bg-emerald-950/20 ring-1 ring-emerald-500/40 shadow-lg shadow-emerald-500/10";
    }

    return `
      <div id="card-${ps.code}" class="rounded-2xl border ${cardBorder} p-5 flex flex-col justify-between relative group">
        
        <div>
          <!-- Top Tag Bar: PS Code, SIH ID, Category, Status -->
          <div class="flex items-center justify-between gap-2 mb-3">
            <div class="flex items-center gap-1.5 flex-wrap">
              <span class="px-2.5 py-1 rounded-lg bg-slate-800 border border-slate-700 text-white font-mono font-extrabold text-xs tracking-wider">
                ${ps.code}
              </span>
              <span class="px-2 py-0.5 rounded bg-slate-800/80 text-slate-400 font-mono text-[10px]">
                ${ps.sihNo}
              </span>
              <span class="px-2 py-0.5 rounded-full border text-[10px] font-semibold flex items-center gap-1 ${categoryBadgeStyle}">
                <i data-lucide="${categoryIcon}" class="w-3 h-3"></i>
                <span>${ps.category}</span>
              </span>
            </div>

            <!-- Availability Status Pill -->
            <div>
              ${isClaimed ? `
                <span class="px-2.5 py-1 rounded-full bg-rose-950/80 text-rose-300 border border-rose-800/60 text-[11px] font-bold flex items-center gap-1">
                  <i data-lucide="lock" class="w-3 h-3"></i>
                  <span>Taken</span>
                </span>
              ` : `
                <span class="px-2.5 py-1 rounded-full bg-emerald-950/80 text-emerald-300 border border-emerald-800/60 text-[11px] font-bold flex items-center gap-1">
                  <span class="w-1.5 h-1.5 rounded-full bg-emerald-400"></span>
                  <span>Available</span>
                </span>
              `}
            </div>
          </div>

          <!-- Domain & Theme Tags -->
          <div class="flex flex-wrap items-center gap-1.5 mb-2.5">
            <span class="px-2 py-0.5 rounded text-[10px] font-semibold border ${domainBadgeStyle}">
              ${ps.domain}
            </span>
            <span class="px-2 py-0.5 rounded text-[10px] font-medium bg-slate-800 text-slate-300 truncate max-w-[200px]" title="${ps.theme}">
              ${ps.theme}
            </span>
          </div>

          <!-- Problem Title -->
          <h3 class="text-sm sm:text-base font-bold text-white group-hover:text-blue-400 transition mb-2 leading-snug">
            ${ps.title}
          </h3>

          <!-- Ministry / Organization -->
          <div class="text-[11px] text-slate-400 mb-3 flex items-center gap-1.5 font-medium">
            <i data-lucide="landmark" class="w-3.5 h-3.5 text-slate-500 flex-shrink-0"></i>
            <span class="truncate">${ps.org}</span>
          </div>

          <!-- Truncated Description -->
          <p class="text-xs text-slate-400 line-clamp-3 leading-relaxed mb-4">
            ${ps.description}
          </p>
        </div>

        <!-- Footer Actions -->
        <div class="pt-3 border-t border-slate-800/80 flex flex-col gap-2">
          
          ${isClaimed ? `
            <!-- Taken State Display -->
            <div class="p-2.5 rounded-xl bg-slate-950/80 border border-slate-800 flex items-center justify-between text-xs">
              <div class="truncate">
                <span class="text-slate-500 text-[10px] block uppercase font-bold">Claimed By:</span>
                <span class="font-bold text-rose-300 truncate">${ps.claimedBy || 'Another Team'}</span>
              </div>
              <button onclick="openDetailModal('${ps.code}')" class="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium">
                View
              </button>
            </div>
          ` : `
            <!-- Available State Actions -->
            <div class="flex items-center gap-2">
              <button onclick="openDetailModal('${ps.code}')" class="px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white text-xs font-semibold transition flex items-center justify-center gap-1">
                <i data-lucide="eye" class="w-3.5 h-3.5"></i>
                <span>Details</span>
              </button>
              
              <button onclick="initiateClaim('${ps.code}')" class="flex-1 py-2 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs shadow-md shadow-blue-500/20 transition flex items-center justify-center gap-1.5 group-hover:bg-blue-500">
                <i data-lucide="check-circle-2" class="w-3.5 h-3.5"></i>
                <span>Select this PS</span>
              </button>
            </div>
          `}

        </div>

      </div>
    `;
  }).join('');

  if (window.lucide) {
    lucide.createIcons();
  }
}

/**
 * Populate Custom Dropdown / Rich Combobox Menu
 * Grouped by Domain
 */
function populateCustomDropdown() {
  const container = document.getElementById('dropdownItemsContainer');
  if (!container) return;

  const domains = ['CSE / AIML', 'Electronics and Communication', 'Mechanical and Civil'];

  let html = '';
  domains.forEach(domain => {
    const items = problemStatements.filter(p => p.domain === domain);
    if (items.length === 0) return;

    html += `
      <div class="bg-slate-950/90 px-3 py-1.5 text-[10px] font-extrabold uppercase tracking-wider text-blue-400 sticky top-0 border-y border-slate-800">
        ${domain} (${items.length})
      </div>
    `;

    items.forEach(ps => {
      const isClaimed = ps.status === 'Claimed' || !!ps.claimedBy;
      const statusPill = isClaimed 
        ? `<span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-rose-950 text-rose-300 border border-rose-900">Taken</span>`
        : `<span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-emerald-950 text-emerald-300 border border-emerald-900">Available</span>`;

      html += `
        <div onclick="selectFromDropdown('${ps.code}')" class="dropdown-item px-3 py-2.5 hover:bg-slate-800/80 cursor-pointer flex items-center justify-between gap-2 transition" data-code="${ps.code}" data-search="${ps.code} ${ps.sihNo} ${ps.title}">
          <div class="truncate">
            <div class="flex items-center gap-1.5">
              <span class="font-mono font-bold text-white text-xs">${ps.code}</span>
              <span class="text-slate-500 text-[10px] font-mono">(${ps.sihNo})</span>
              <span class="text-[10px] text-slate-400">&bull; ${ps.category}</span>
            </div>
            <div class="text-[11px] text-slate-300 truncate font-medium mt-0.5">${ps.title}</div>
          </div>
          <div>${statusPill}</div>
        </div>
      `;
    });
  });

  container.innerHTML = html;
}

/**
 * Filter items inside the custom dropdown popup
 */
function filterDropdownList(query) {
  const q = (query || '').toLowerCase().trim();
  const items = document.querySelectorAll('.dropdown-item');
  items.forEach(item => {
    const text = item.getAttribute('data-search').toLowerCase();
    if (text.includes(q)) {
      item.classList.remove('hidden');
    } else {
      item.classList.add('hidden');
    }
  });
}

function toggleCustomDropdown() {
  const menu = document.getElementById('customDropdownMenu');
  const arrow = document.getElementById('customDropdownArrow');
  if (menu) {
    const isHidden = menu.classList.toggle('hidden');
    if (arrow) {
      if (!isHidden) {
        arrow.classList.add('rotate-180');
        const input = document.getElementById('dropdownSearchInput');
        if (input) setTimeout(() => input.focus(), 50);
      } else {
        arrow.classList.remove('rotate-180');
      }
    }
  }
}

function selectFromDropdown(code) {
  toggleCustomDropdown();
  const label = document.getElementById('customDropdownLabel');
  if (label) {
    label.textContent = `Selected: ${code}`;
  }

  // Scroll to card smoothly
  const card = document.getElementById(`card-${code}`);
  if (card) {
    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    card.classList.add('ring-2', 'ring-blue-500');
    setTimeout(() => card.classList.remove('ring-2', 'ring-blue-500'), 2500);
  }

  // Also open detail modal
  openDetailModal(code);
}

/**
 * Update Header and Filter Counter Badges
 */
function updateCounts() {
  let total = problemStatements.length;
  let available = 0;
  let claimed = 0;

  let cseCount = 0;
  let eceCount = 0;
  let mechCivCount = 0;

  problemStatements.forEach(ps => {
    const isClaimed = ps.status === 'Claimed' || !!ps.claimedBy;
    if (isClaimed) claimed++;
    else available++;

    if (ps.domain === 'CSE / AIML') cseCount++;
    else if (ps.domain === 'Electronics and Communication') eceCount++;
    else if (ps.domain === 'Mechanical and Civil') mechCivCount++;
  });

  const statAvail = document.getElementById('statAvailable');
  const statClaim = document.getElementById('statClaimed');
  if (statAvail) statAvail.textContent = available;
  if (statClaim) statClaim.textContent = claimed;

  const countAll = document.getElementById('countDomainAll');
  const countCSE = document.getElementById('countDomainCSE');
  const countECE = document.getElementById('countDomainECE');
  const countMC = document.getElementById('countDomainMechCiv');

  if (countAll) countAll.textContent = total;
  if (countCSE) countCSE.textContent = cseCount;
  if (countECE) countECE.textContent = eceCount;
  if (countMC) countMC.textContent = mechCivCount;
}

/**
 * Open Problem Statement Details Modal
 */
function openDetailModal(psCode) {
  const ps = problemStatements.find(p => p.code === psCode);
  if (!ps) return;

  const modal = document.getElementById('detailModal');
  const isClaimed = ps.status === 'Claimed' || !!ps.claimedBy;

  document.getElementById('modalPsCode').textContent = ps.code;
  document.getElementById('modalSihNo').textContent = ps.sihNo;
  document.getElementById('modalTitle').textContent = ps.title;
  document.getElementById('modalDomain').textContent = ps.domain;
  document.getElementById('modalTheme').textContent = ps.theme;
  document.getElementById('modalOrg').textContent = ps.org;
  document.getElementById('modalDescription').textContent = ps.description;

  const catBadge = document.getElementById('modalCategoryBadge');
  if (catBadge) {
    catBadge.textContent = ps.category;
    catBadge.className = ps.category === 'Software'
      ? 'text-[11px] font-semibold px-2 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30'
      : 'text-[11px] font-semibold px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30';
  }

  const selectBtn = document.getElementById('modalSelectBtn');
  if (selectBtn) {
    if (isClaimed) {
      selectBtn.disabled = true;
      selectBtn.className = 'px-5 py-2 rounded-xl bg-slate-800 text-slate-500 font-bold text-xs cursor-not-allowed';
      selectBtn.innerHTML = `<span>🔒 Already Claimed</span>`;
    } else {
      selectBtn.disabled = false;
      selectBtn.className = 'px-5 py-2 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs shadow-lg shadow-blue-500/20 transition flex items-center gap-1.5';
      selectBtn.innerHTML = `<i data-lucide="check" class="w-4 h-4"></i><span>Select this PS for Team</span>`;
    }
  }

  modal.setAttribute('data-current-ps', ps.code);
  modal.classList.remove('hidden');

  if (window.lucide) {
    lucide.createIcons();
  }
}

function closeDetailModal() {
  const modal = document.getElementById('detailModal');
  if (modal) modal.classList.add('hidden');
}

function confirmSelectFromModal() {
  const modal = document.getElementById('detailModal');
  const psCode = modal.getAttribute('data-current-ps');
  closeDetailModal();
  if (psCode) {
    initiateClaim(psCode);
  }
}

/**
 * Step 1: Initiate Problem Statement Claim Flow
 * Validates team information and prompts confirmation modal
 */
function initiateClaim(psCode) {
  const teamName = (document.getElementById('inputTeamName').value || '').trim();
  const dept = (document.getElementById('inputDepartment').value || '').trim();
  const leaderName = (document.getElementById('inputLeaderName').value || '').trim();
  const leaderContact = (document.getElementById('inputLeaderContact').value || '').trim();

  // Validate Team Name
  if (!teamName) {
    showToast("Please enter your Team Name in Step 1 first!", "error");
    document.getElementById('inputTeamName').focus();
    return;
  }

  // Validate Leader Name
  if (!leaderName) {
    showToast("Please enter your Team Leader Name!", "error");
    document.getElementById('inputLeaderName').focus();
    return;
  }

  // Validate Contact
  if (!leaderContact) {
    showToast("Please enter Leader Contact / WhatsApp number!", "error");
    document.getElementById('inputLeaderContact').focus();
    return;
  }

  // Check if team already holds a PS in this browser
  if (userClaimedPS) {
    showToast(`Your team already claimed PS [${userClaimedPS.psCode}]. Multiple selections are not permitted.`, "error");
    return;
  }

  const ps = problemStatements.find(p => p.code === psCode);
  if (!ps) return;

  if (ps.status === 'Claimed' || !!ps.claimedBy) {
    showToast(`Sorry, ${ps.code} has already been claimed!`, "error");
    return;
  }

  // Save entered team details to localStorage
  saveTeamInfo(teamName, dept, leaderName, leaderContact);

  // Setup Confirmation Modal
  pendingClaimPS = ps;
  document.getElementById('confirmTeam').textContent = teamName;
  document.getElementById('confirmDept').textContent = dept;
  document.getElementById('confirmPsCode').textContent = `${ps.code} (${ps.sihNo})`;
  document.getElementById('confirmPsTitle').textContent = ps.title;

  document.getElementById('confirmModal').classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function closeConfirmModal() {
  document.getElementById('confirmModal').classList.add('hidden');
  pendingClaimPS = null;
}

/**
 * Step 2: Execute Atomic Concurrency-Safe Claim
 * Calls Google Apps Script Web App with ScriptLock
 */
async function executeClaim() {
  if (!pendingClaimPS) return;

  const ps = pendingClaimPS;
  const teamName = (document.getElementById('inputTeamName').value || '').trim();
  const dept = (document.getElementById('inputDepartment').value || '').trim();
  const leaderName = (document.getElementById('inputLeaderName').value || '').trim();
  const leaderContact = (document.getElementById('inputLeaderContact').value || '').trim();

  const submitBtn = document.getElementById('finalSubmitBtn');
  const btnSpinner = document.getElementById('btnSpinner');
  const btnText = document.getElementById('btnText');

  // Show loading spinner
  submitBtn.disabled = true;
  if (btnSpinner) btnSpinner.classList.remove('hidden');
  if (btnText) btnText.textContent = "Locking in Sheet...";

  const payload = {
    action: 'claim',
    psCode: ps.code,
    teamName: teamName,
    department: dept,
    leaderName: leaderName,
    leaderContact: leaderContact,
    timestamp: new Date().toISOString()
  };

  // If Google Apps Script URL is configured, call it
  if (appsScriptUrl) {
    try {
      let result = null;

      // Primary Attempt: POST with text/plain body to avoid CORS preflight issues
      try {
        const postRes = await fetch(appsScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload)
        });
        result = await postRes.json();
      } catch (postErr) {
        console.warn("POST attempt had network issue, trying GET claim fallback...", postErr);
        // Fallback Attempt: GET claim query param
        const getUrl = `${appsScriptUrl}?action=claim&teamName=${encodeURIComponent(teamName)}&department=${encodeURIComponent(dept)}&leaderName=${encodeURIComponent(leaderName)}&leaderContact=${encodeURIComponent(leaderContact)}&psCode=${encodeURIComponent(ps.code)}&_t=${Date.now()}`;
        const getRes = await fetch(getUrl, { method: 'GET', mode: 'cors' });
        result = await getRes.json();
      }

      if (result && result.success) {
        handleSuccessfulClaim(ps, teamName, dept, leaderName, leaderContact, result.data?.timestamp);
      } else {
        const msg = result ? result.message : "Selection could not be locked.";
        showToast(`Selection Failed: ${msg}`, "error");
        // Re-sync immediately to get who actually claimed it
        refreshData(false);
      }
    } catch (err) {
      console.error("Claim error:", err);
      showToast("Network error while contacting Google Sheets. Please verify your connection or try again.", "error");
      refreshData(false);
    } finally {
      submitBtn.disabled = false;
      if (btnSpinner) btnSpinner.classList.add('hidden');
      if (btnText) btnText.textContent = "Confirm & Lock PS";
      closeConfirmModal();
    }
  } else {
    // Demo Mode Claim simulation
    setTimeout(() => {
      // Mock Claim in localStorage
      const mockClaims = JSON.parse(localStorage.getItem('ivc_mock_claims') || '{}');
      if (mockClaims[ps.code]) {
        showToast(`Sorry, ${ps.code} was claimed in Demo Mode!`, "error");
        closeConfirmModal();
        submitBtn.disabled = false;
        if (btnSpinner) btnSpinner.classList.add('hidden');
        if (btnText) btnText.textContent = "Confirm & Lock PS";
        return;
      }

      const timestamp = new Date().toLocaleString();
      mockClaims[ps.code] = { teamName, dept, leaderName, leaderContact, timestamp };
      localStorage.setItem('ivc_mock_claims', JSON.stringify(mockClaims));

      handleSuccessfulClaim(ps, teamName, dept, leaderName, leaderContact, timestamp);

      submitBtn.disabled = false;
      if (btnSpinner) btnSpinner.classList.add('hidden');
      if (btnText) btnText.textContent = "Confirm & Lock PS";
      closeConfirmModal();
    }, 800);
  }
}

/**
 * Handle Successful Claim Execution
 */
function handleSuccessfulClaim(ps, teamName, dept, leaderName, leaderContact, timestamp) {
  const tsStr = timestamp || new Date().toLocaleString();

  // 1. Update In-Memory State
  ps.status = 'Claimed';
  ps.claimedBy = teamName;
  ps.department = dept;
  ps.claimedAt = tsStr;

  // 2. Persist to localStorage
  userClaimedPS = {
    psCode: ps.code,
    sihNo: ps.sihNo,
    title: ps.title,
    domain: ps.domain,
    category: ps.category,
    teamName: teamName,
    department: dept,
    leaderName: leaderName,
    leaderContact: leaderContact,
    timestamp: tsStr
  };
  localStorage.setItem('ivc_user_claimed_ps', JSON.stringify(userClaimedPS));

  // 3. Update UI
  updateClaimedBanner();
  renderProblemStatements();
  updateCounts();
  populateCustomDropdown();

  // 4. Show Celebratory Receipt Modal
  showClaimReceiptModal();
}

/**
 * Display User's Confirmed Selection Banner & Lock Team Info
 */
function updateClaimedBanner() {
  const banner = document.getElementById('activeClaimBanner');
  const lockNotice = document.getElementById('teamLockedNotice');

  if (userClaimedPS) {
    if (banner) {
      document.getElementById('claimedBannerTitle').textContent = `[${userClaimedPS.psCode}] ${userClaimedPS.title}`;
      document.getElementById('claimedBannerTeam').textContent = userClaimedPS.teamName;
      document.getElementById('claimedBannerTime').textContent = userClaimedPS.timestamp;
      banner.classList.remove('hidden');
    }
    if (lockNotice) {
      lockNotice.classList.remove('hidden');
    }

    // Disable form edits once claimed
    ['inputTeamName', 'inputDepartment', 'inputLeaderName', 'inputLeaderContact'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.disabled = true;
    });
  } else {
    if (banner) banner.classList.add('hidden');
    if (lockNotice) lockNotice.classList.add('hidden');
  }
}

/**
 * Show Claim Receipt Ticket Modal
 */
function showClaimReceiptModal() {
  if (!userClaimedPS) return;

  document.getElementById('receiptTeamName').textContent = userClaimedPS.teamName;
  document.getElementById('receiptDept').textContent = userClaimedPS.department;
  document.getElementById('receiptLeader').textContent = `${userClaimedPS.leaderName} (${userClaimedPS.leaderContact})`;
  document.getElementById('receiptPsCode').textContent = userClaimedPS.psCode;
  document.getElementById('receiptSihNo').textContent = userClaimedPS.sihNo;
  document.getElementById('receiptTitle').textContent = userClaimedPS.title;
  document.getElementById('receiptTimestamp').textContent = userClaimedPS.timestamp;

  document.getElementById('receiptModal').classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function closeReceiptModal() {
  document.getElementById('receiptModal').classList.add('hidden');
}

/**
 * Team Info LocalStorage Helpers
 */
function saveTeamInfo(teamName, dept, leaderName, leaderContact) {
  try {
    localStorage.setItem('ivc_team_info', JSON.stringify({ teamName, dept, leaderName, leaderContact }));
  } catch (e) {}
}

function loadSavedTeamInfo() {
  try {
    const saved = localStorage.getItem('ivc_team_info');
    if (saved) {
      const info = JSON.parse(saved);
      if (info.teamName) document.getElementById('inputTeamName').value = info.teamName;
      if (info.dept) document.getElementById('inputDepartment').value = info.dept;
      if (info.leaderName) document.getElementById('inputLeaderName').value = info.leaderName;
      if (info.leaderContact) document.getElementById('inputLeaderContact').value = info.leaderContact;
    }
  } catch (e) {}
}

/**
 * Reset all filters to default
 */
function resetFilters() {
  currentDomainFilter = 'ALL';
  setDomainFilter('ALL');
  document.getElementById('categorySelect').value = 'ALL';
  document.getElementById('statusSelect').value = 'ALL';
  document.getElementById('searchInput').value = '';
  renderProblemStatements();
}

/**
 * Google Apps Script Web App Configuration Modal
 */
function openConfigModal() {
  const modal = document.getElementById('configModal');
  const input = document.getElementById('configApiUrl');
  if (input) input.value = appsScriptUrl;
  if (modal) modal.classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function closeConfigModal() {
  const modal = document.getElementById('configModal');
  if (modal) modal.classList.add('hidden');
}

function saveConfigApiUrl() {
  const input = document.getElementById('configApiUrl');
  const url = (input.value || '').trim();
  appsScriptUrl = url;
  localStorage.setItem('ivc_apps_script_url', url);

  closeConfigModal();
  showToast("Google Sheets Web App URL Saved!", "success");
  startLiveSync();
}

async function testConnection() {
  const input = document.getElementById('configApiUrl');
  const url = (input.value || '').trim();
  const resultDiv = document.getElementById('testConnResult');
  const testBtn = document.getElementById('testConnBtn');

  if (!url) {
    resultDiv.className = 'text-xs p-2.5 rounded-lg bg-rose-950/80 text-rose-300 border border-rose-800';
    resultDiv.textContent = 'Please enter a valid Google Apps Script Web App URL first.';
    resultDiv.classList.remove('hidden');
    return;
  }

  testBtn.disabled = true;
  testBtn.innerHTML = `<i data-lucide="loader-2" class="w-3.5 h-3.5 animate-spin"></i><span>Testing...</span>`;
  if (window.lucide) lucide.createIcons();

  try {
    const testUrl = url.includes('?') ? `${url}&action=ping&_t=${Date.now()}` : `${url}?action=ping&_t=${Date.now()}`;
    const res = await fetch(testUrl, { method: 'GET', mode: 'cors' });
    const json = await res.json();

    resultDiv.className = 'text-xs p-2.5 rounded-lg bg-emerald-950/80 text-emerald-300 border border-emerald-800 flex items-center gap-1.5';
    resultDiv.innerHTML = `<i data-lucide="check" class="w-4 h-4 text-emerald-400"></i><span>Connected successfully! (${json.message || 'Service is online'})</span>`;
    resultDiv.classList.remove('hidden');
  } catch (err) {
    resultDiv.className = 'text-xs p-2.5 rounded-lg bg-rose-950/80 text-rose-300 border border-rose-800';
    resultDiv.textContent = `Connection failed: ${err.message}. Make sure the script is deployed as Web App with 'Who has access: Anyone'.`;
    resultDiv.classList.remove('hidden');
  } finally {
    testBtn.disabled = false;
    testBtn.innerHTML = `<i data-lucide="wifi" class="w-3.5 h-3.5"></i><span>Test Connection</span>`;
    if (window.lucide) lucide.createIcons();
  }
}

/**
 * Toast Notifications
 */
let toastTimeout = null;
function showToast(message, type = "info") {
  const toast = document.getElementById('toastNotification');
  const toastInner = document.getElementById('toastInner');
  const toastMsg = document.getElementById('toastMsg');
  const toastIcon = document.getElementById('toastIcon');

  if (!toast || !toastMsg) return;

  if (toastTimeout) clearTimeout(toastTimeout);

  toastMsg.textContent = message;

  if (type === 'error') {
    toastInner.className = 'p-3.5 rounded-xl shadow-2xl border text-xs flex items-center justify-between gap-3 bg-rose-950 text-rose-200 border-rose-800';
    toastIcon.setAttribute('data-lucide', 'alert-circle');
  } else if (type === 'success') {
    toastInner.className = 'p-3.5 rounded-xl shadow-2xl border text-xs flex items-center justify-between gap-3 bg-emerald-950 text-emerald-200 border-emerald-800';
    toastIcon.setAttribute('data-lucide', 'check-circle');
  } else {
    toastInner.className = 'p-3.5 rounded-xl shadow-2xl border text-xs flex items-center justify-between gap-3 bg-slate-900 text-slate-200 border-slate-700';
    toastIcon.setAttribute('data-lucide', 'info');
  }

  if (window.lucide) lucide.createIcons();

  toast.classList.remove('translate-y-24', 'opacity-0');
  toast.classList.add('translate-y-0', 'opacity-100');

  toastTimeout = setTimeout(() => {
    hideToast();
  }, 4000);
}

function hideToast() {
  const toast = document.getElementById('toastNotification');
  if (toast) {
    toast.classList.remove('translate-y-0', 'opacity-100');
    toast.classList.add('translate-y-24', 'opacity-0');
  }
}
