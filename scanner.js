const { execSync, exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const os = require("os");
const initSqlJs = require("sql.js");

function runPS(script, timeout = 60000) {
  try {
    const result = execSync(
      `powershell -NoProfile -ExecutionPolicy Bypass -Command "${script.replace(/"/g, '\\"')}"`,
      { encoding: "utf-8", maxBuffer: 50 * 1024 * 1024, timeout, windowsHide: true }
    );
    return result.trim();
  } catch (e) {
    console.error(`PS error: ${e.message?.slice(0, 200)}`);
    return "";
  }
}

function parseJsonOutput(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}

// 1. Installed Programs (Registry - both 64-bit and 32-bit)
function getInstalledPrograms() {
  console.log("  Scanning installed programs (registry)...");
  const ps = `
    $paths = @(
      'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',
      'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',
      'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'
    )
    $apps = $paths | ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
      Where-Object { $_.DisplayName -and $_.DisplayName -ne '' } |
      Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, EstimatedSize, URLInfoAbout |
      Sort-Object DisplayName -Unique
    $apps | ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps));
}

// 2. Running Processes with details
function getRunningProcesses() {
  console.log("  Scanning running processes...");
  const ps = `
    Get-Process | Where-Object { $_.MainWindowTitle -or $_.Path } |
      Select-Object Name, @{N='PID';E={$_.Id}},
        @{N='MemoryMB';E={[math]::Round($_.WorkingSet64/1MB,1)}},
        @{N='CPU_Seconds';E={[math]::Round($_.CPU,1)}},
        Path, Description, Company, MainWindowTitle,
        @{N='StartTime';E={try{$_.StartTime.ToString('yyyy-MM-dd HH:mm:ss')}catch{''}}} |
      Sort-Object MemoryMB -Descending |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps));
}

// 3. Startup Programs
function getStartupPrograms() {
  console.log("  Scanning startup programs...");
  const ps = `
    $startup = @()
    # Registry Run keys
    $regPaths = @(
      'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run',
      'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run',
      'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce',
      'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce'
    )
    foreach ($p in $regPaths) {
      $items = Get-ItemProperty $p -ErrorAction SilentlyContinue
      if ($items) {
        $items.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
          $startup += [PSCustomObject]@{ Name=$_.Name; Command=$_.Value; Source=$p }
        }
      }
    }
    # Startup folder
    $startupFolder = [Environment]::GetFolderPath('Startup')
    Get-ChildItem $startupFolder -ErrorAction SilentlyContinue | ForEach-Object {
      $startup += [PSCustomObject]@{ Name=$_.BaseName; Command=$_.FullName; Source='StartupFolder' }
    }
    # Scheduled tasks that run at logon
    Get-ScheduledTask -ErrorAction SilentlyContinue |
      Where-Object { $_.Triggers -and $_.State -ne 'Disabled' } |
      ForEach-Object {
        $hasLogon = $_.Triggers | Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskLogonTrigger' }
        if ($hasLogon) {
          $startup += [PSCustomObject]@{ Name=$_.TaskName; Command=$_.Actions.Execute; Source='ScheduledTask-Logon' }
        }
      }
    $startup | ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 90000));
}

// 4. Browser History - Chrome
async function getChromiumHistory(browserName, profileBase, SQL) {
  console.log(`  Scanning ${browserName} history...`);
  if (!fs.existsSync(profileBase)) return [];

  // Find all profile dirs
  const profiles = [];
  try {
    const entries = fs.readdirSync(profileBase);
    for (const e of entries) {
      const histPath = path.join(profileBase, e, "History");
      if (fs.existsSync(histPath)) profiles.push({ profile: e, path: histPath });
    }
    // Also check root
    const rootHist = path.join(profileBase, "History");
    if (fs.existsSync(rootHist)) profiles.push({ profile: "Default", path: rootHist });
  } catch { }

  const allHistory = [];
  for (const p of profiles) {
    const tmp = path.join(os.tmpdir(), `${browserName}_${p.profile}_hist_${Date.now()}.db`);
    try {
      fs.copyFileSync(p.path, tmp);
      const buffer = fs.readFileSync(tmp);
      const db = new SQL.Database(buffer);
      try {
        const results = db.exec(`
          SELECT url, title, visit_count,
            datetime(last_visit_time/1000000-11644473600, 'unixepoch', 'localtime') as last_visit
          FROM urls
          WHERE last_visit_time > 0
          ORDER BY visit_count DESC
          LIMIT 5000
        `);
        if (results.length > 0) {
          const cols = results[0].columns;
          for (const row of results[0].values) {
            const obj = {};
            cols.forEach((c, i) => obj[c] = row[i]);
            allHistory.push({ ...obj, browser: browserName, profile: p.profile });
          }
        }
      } finally {
        db.close();
      }
    } catch (e) {
      console.error(`  ${browserName} history read error: ${e.message?.slice(0, 100)}`);
    } finally {
      try { fs.unlinkSync(tmp); } catch { }
    }
  }
  return allHistory;
}

async function getBrowserHistory(SQL) {
  const userDir = os.homedir();
  const results = [];

  // Chrome
  results.push(...await getChromiumHistory("Chrome",
    path.join(userDir, "AppData", "Local", "Google", "Chrome", "User Data"), SQL));

  // Edge
  results.push(...await getChromiumHistory("Edge",
    path.join(userDir, "AppData", "Local", "Microsoft", "Edge", "User Data"), SQL));

  // Brave
  results.push(...await getChromiumHistory("Brave",
    path.join(userDir, "AppData", "Local", "BraveSoftware", "Brave-Browser", "User Data"), SQL));

  // Firefox (different DB schema)
  console.log("  Scanning Firefox history...");
  const ffBase = path.join(userDir, "AppData", "Roaming", "Mozilla", "Firefox", "Profiles");
  if (fs.existsSync(ffBase)) {
    try {
      const profiles = fs.readdirSync(ffBase);
      for (const prof of profiles) {
        const placesPath = path.join(ffBase, prof, "places.sqlite");
        if (!fs.existsSync(placesPath)) continue;
        const tmp = path.join(os.tmpdir(), `ff_${prof}_${Date.now()}.db`);
        try {
          fs.copyFileSync(placesPath, tmp);
          const buffer = fs.readFileSync(tmp);
          const db = new SQL.Database(buffer);
          try {
            const res = db.exec(`
              SELECT url, title, visit_count,
                datetime(last_visit_date/1000000, 'unixepoch', 'localtime') as last_visit
              FROM moz_places
              WHERE visit_count > 0
              ORDER BY visit_count DESC
              LIMIT 5000
            `);
            if (res.length > 0) {
              const cols = res[0].columns;
              for (const row of res[0].values) {
                const obj = {};
                cols.forEach((c, i) => obj[c] = row[i]);
                results.push({ ...obj, browser: "Firefox", profile: prof });
              }
            }
          } finally {
            db.close();
          }
        } catch (e) {
          console.error(`  Firefox history error: ${e.message?.slice(0, 100)}`);
        } finally {
          try { fs.unlinkSync(tmp); } catch { }
        }
      }
    } catch { }
  }

  return results;
}

// 5. Windows Prefetch (app launch frequency)
function getPrefetchData() {
  console.log("  Scanning Windows Prefetch data...");
  const ps = `
    $prefetchDir = "$env:SystemRoot\\Prefetch"
    if (Test-Path $prefetchDir) {
      Get-ChildItem "$prefetchDir\\*.pf" -ErrorAction SilentlyContinue |
        Select-Object @{N='Name';E={$_.BaseName -replace '-[A-F0-9]+$',''}},
          @{N='LastRun';E={$_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')}},
          @{N='Created';E={$_.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')}},
          @{N='RunCount';E={
            try {
              $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
              if ($bytes.Length -gt 208) {
                [BitConverter]::ToInt32($bytes, 208)
              } else { -1 }
            } catch { -1 }
          }} |
        Sort-Object LastRun -Descending |
        ConvertTo-Json -Depth 3
    } else { '[]' }
  `;
  return parseJsonOutput(runPS(ps, 90000));
}

// 6. Recently used files (Shell:Recent)
function getRecentFiles() {
  console.log("  Scanning recent files...");
  const ps = `
    $recent = [Environment]::GetFolderPath('Recent')
    Get-ChildItem $recent -Filter *.lnk -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 500 @{N='Name';E={$_.BaseName}},
        @{N='LastAccessed';E={$_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')}},
        @{N='Target';E={
          try {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($_.FullName)
            $shortcut.TargetPath
          } catch { '' }
        }} |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 90000));
}

// 7. Windows Store / UWP Apps
function getStoreApps() {
  console.log("  Scanning Windows Store apps...");
  const ps = `
    Get-AppxPackage -ErrorAction SilentlyContinue |
      Where-Object { $_.IsFramework -eq $false -and $_.SignatureKind -eq 'Store' } |
      Select-Object Name, @{N='DisplayName';E={
        try {
          (Get-AppxPackageManifest $_).Package.Properties.DisplayName
        } catch { $_.Name }
      }}, Version, Publisher, InstallLocation,
        @{N='InstalledDate';E={
          try { (Get-Item $_.InstallLocation).CreationTime.ToString('yyyy-MM-dd') } catch { '' }
        }} |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 120000));
}

// 8. Services
function getServices() {
  console.log("  Scanning Windows services...");
  const ps = `
    Get-Service | Where-Object { $_.Status -eq 'Running' } |
      Select-Object Name, DisplayName, Status,
        @{N='StartType';E={$_.StartType.ToString()}},
        @{N='Description';E={
          try { (Get-WmiObject Win32_Service -Filter "Name='$($_.Name)'" -ErrorAction SilentlyContinue).Description } catch { '' }
        }} |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 120000));
}

// 9. Scheduled Tasks
function getScheduledTasks() {
  console.log("  Scanning scheduled tasks...");
  const ps = `
    Get-ScheduledTask -ErrorAction SilentlyContinue |
      Where-Object { $_.State -ne 'Disabled' -and $_.TaskPath -notlike '\\Microsoft\\*' } |
      Select-Object TaskName, TaskPath, State,
        @{N='Action';E={($_.Actions | Select-Object -First 1).Execute}},
        @{N='Arguments';E={($_.Actions | Select-Object -First 1).Arguments}},
        @{N='Author';E={$_.Author}} |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 90000));
}

// 10. Network connections (detect apps phoning home)
function getNetworkConnections() {
  console.log("  Scanning active network connections...");
  const ps = `
    Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue |
      ForEach-Object {
        $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
          ProcessName = $proc.Name
          ProcessPath = $proc.Path
          RemoteAddress = $_.RemoteAddress
          RemotePort = $_.RemotePort
          LocalPort = $_.LocalPort
        }
      } |
      Where-Object { $_.RemoteAddress -ne '127.0.0.1' -and $_.RemoteAddress -ne '::1' } |
      ConvertTo-Json -Depth 3
  `;
  return parseJsonOutput(runPS(ps, 60000));
}

// Aggregate SaaS domains from browser history
function extractSaasFromHistory(history) {
  const domainMap = {};
  for (const entry of history) {
    try {
      const url = new URL(entry.url);
      const domain = url.hostname.replace(/^www\./, "");
      if (!domainMap[domain]) {
        domainMap[domain] = { domain, totalVisits: 0, titles: new Set(), browsers: new Set(), lastVisit: "" };
      }
      domainMap[domain].totalVisits += entry.visit_count || 1;
      if (entry.title) domainMap[domain].titles.add(entry.title);
      if (entry.browser) domainMap[domain].browsers.add(entry.browser);
      if (entry.last_visit && entry.last_visit > domainMap[domain].lastVisit) {
        domainMap[domain].lastVisit = entry.last_visit;
      }
    } catch { }
  }

  // Convert sets to arrays and sort by visits
  return Object.values(domainMap)
    .map(d => ({
      ...d,
      titles: [...d.titles].slice(0, 3),
      browsers: [...d.browsers]
    }))
    .sort((a, b) => b.totalVisits - a.totalVisits);
}

// Known SaaS patterns for categorization
const SAAS_PATTERNS = {
  "Productivity & Docs": [
    "docs.google.com", "sheets.google.com", "slides.google.com", "drive.google.com",
    "notion.so", "notion.site", "airtable.com", "coda.io", "clickup.com",
    "monday.com", "asana.com", "trello.com", "basecamp.com", "todoist.com",
    "evernote.com", "onenote.com", "dropbox.com", "box.com",
    "office.com", "office365.com", "sharepoint.com", "onedrive.live.com",
    "quip.com", "roamresearch.com", "obsidian.md"
  ],
  "Communication & Email": [
    "mail.google.com", "outlook.live.com", "outlook.office.com", "outlook.office365.com",
    "slack.com", "app.slack.com", "teams.microsoft.com", "discord.com", "discord.gg",
    "zoom.us", "meet.google.com", "whereby.com", "loom.com", "calendly.com",
    "telegram.org", "web.telegram.org", "web.whatsapp.com", "messages.google.com",
    "intercom.com", "crisp.chat", "zendesk.com", "freshdesk.com"
  ],
  "Design & Creative": [
    "figma.com", "canva.com", "miro.com", "whimsical.com", "lucidchart.com",
    "sketch.com", "invisionapp.com", "adobe.com", "creativecloud.adobe.com",
    "photopea.com", "pixlr.com", "remove.bg", "coolors.co", "dribbble.com",
    "behance.net", "unsplash.com", "pexels.com"
  ],
  "Development & DevOps": [
    "github.com", "gitlab.com", "bitbucket.org", "stackoverflow.com",
    "vercel.com", "netlify.com", "heroku.com", "railway.app", "render.com",
    "docker.com", "hub.docker.com", "aws.amazon.com", "console.aws.amazon.com",
    "cloud.google.com", "console.cloud.google.com", "portal.azure.com",
    "firebase.google.com", "supabase.com", "planetscale.com", "neon.tech",
    "sentry.io", "datadog.com", "grafana.com", "newrelic.com",
    "npmjs.com", "pypi.org", "crates.io", "packagist.org",
    "codepen.io", "codesandbox.io", "replit.com", "stackblitz.com",
    "postman.com", "insomnia.rest", "swagger.io"
  ],
  "AI & ML Tools": [
    "chat.openai.com", "chatgpt.com", "openai.com", "platform.openai.com",
    "claude.ai", "anthropic.com", "console.anthropic.com",
    "bard.google.com", "gemini.google.com", "ai.google.dev",
    "huggingface.co", "replicate.com", "midjourney.com",
    "copilot.github.com", "cursor.com", "cursor.sh",
    "perplexity.ai", "phind.com", "you.com",
    "colab.research.google.com", "kaggle.com",
    "together.ai", "groq.com", "mistral.ai",
    "lovable.dev", "v0.dev", "bolt.new"
  ],
  "Marketing & Analytics": [
    "analytics.google.com", "search.google.com", "ads.google.com",
    "mailchimp.com", "sendgrid.com", "hubspot.com", "salesforce.com",
    "hootsuite.com", "buffer.com", "later.com", "sproutsocial.com",
    "semrush.com", "ahrefs.com", "moz.com", "hotjar.com",
    "mixpanel.com", "amplitude.com", "segment.com", "heap.io",
    "typeform.com", "surveymonkey.com", "google.com/forms",
    "meta.com", "business.facebook.com", "ads.twitter.com"
  ],
  "Finance & Payments": [
    "dashboard.stripe.com", "stripe.com", "paypal.com",
    "razorpay.com", "quickbooks.intuit.com", "xero.com",
    "freshbooks.com", "wave.com", "invoice.zoho.com",
    "wise.com", "revolut.com", "mercury.com"
  ],
  "Learning & Reference": [
    "udemy.com", "coursera.org", "edx.org", "pluralsight.com",
    "linkedin.com/learning", "skillshare.com", "masterclass.com",
    "medium.com", "dev.to", "hashnode.com", "substack.com",
    "youtube.com", "youtu.be"
  ],
  "Project & Knowledge Management": [
    "linear.app", "jira.atlassian.com", "atlassian.com",
    "confluence.atlassian.com", "notion.so", "coda.io",
    "productboard.com", "shortcut.com", "height.app"
  ],
  "Security & Identity": [
    "1password.com", "bitwarden.com", "lastpass.com",
    "okta.com", "auth0.com", "clerk.com"
  ],
  "Social & Networking": [
    "linkedin.com", "twitter.com", "x.com", "facebook.com",
    "instagram.com", "reddit.com", "news.ycombinator.com",
    "producthunt.com", "threads.net", "mastodon.social",
    "bsky.app"
  ]
};

// Desktop app categorization
const DESKTOP_APP_PATTERNS = {
  "Development Tools": [
    "visual studio", "vs code", "vscode", "code.exe", "intellij", "pycharm", "webstorm",
    "android studio", "xcode", "sublime", "atom", "notepad++", "vim", "neovim",
    "git", "github desktop", "sourcetree", "gitkraken", "docker", "postman",
    "wsl", "terminal", "powershell", "windows terminal", "hyper", "iterm",
    "node", "python", "java", "go", "rust", "ruby", "cursor", "windsurf"
  ],
  "Productivity": [
    "microsoft office", "word", "excel", "powerpoint", "outlook", "onenote",
    "libreoffice", "wps office", "google drive", "dropbox", "onedrive",
    "notion", "obsidian", "todoist", "ticktick", "any.do"
  ],
  "Communication": [
    "slack", "teams", "discord", "zoom", "skype", "telegram", "whatsapp",
    "signal", "webex", "google chat", "thunderbird"
  ],
  "Design & Media": [
    "photoshop", "illustrator", "figma", "sketch", "gimp", "inkscape",
    "blender", "after effects", "premiere", "davinci", "obs", "audacity",
    "vlc", "spotify", "itunes", "foobar"
  ],
  "Browsers": [
    "chrome", "firefox", "edge", "brave", "opera", "vivaldi", "arc"
  ],
  "Security & Utilities": [
    "antivirus", "malwarebytes", "norton", "kaspersky", "bitdefender",
    "ccleaner", "7-zip", "winrar", "everything", "autohotkey",
    "1password", "bitwarden", "lastpass", "keepass"
  ],
  "AI Tools": [
    "chatgpt", "claude", "copilot", "cursor", "codeium", "tabnine"
  ]
};

function categorizeApp(name, publisher) {
  const lower = (name + " " + (publisher || "")).toLowerCase();
  for (const [cat, patterns] of Object.entries(DESKTOP_APP_PATTERNS)) {
    for (const p of patterns) {
      if (lower.includes(p.toLowerCase())) return cat;
    }
  }
  return "Other";
}

function categorizeDomain(domain) {
  for (const [cat, domains] of Object.entries(SAAS_PATTERNS)) {
    for (const d of domains) {
      if (domain === d || domain.endsWith("." + d)) return cat;
    }
  }
  return null;
}

// Main scan function
async function runFullScan() {
  console.log("\n=== Starting Full Windows PC Scan ===\n");
  const startTime = Date.now();

  // Initialize sql.js
  const SQL = await initSqlJs();

  const results = {};

  // Run scans
  results.installedPrograms = getInstalledPrograms();
  console.log(`  Found ${results.installedPrograms.length} installed programs`);

  results.runningProcesses = getRunningProcesses();
  console.log(`  Found ${results.runningProcesses.length} running processes`);

  results.startupPrograms = getStartupPrograms();
  console.log(`  Found ${results.startupPrograms.length} startup entries`);

  results.browserHistory = await getBrowserHistory(SQL);
  console.log(`  Found ${results.browserHistory.length} browser history entries`);

  results.prefetchData = getPrefetchData();
  console.log(`  Found ${results.prefetchData.length} prefetch entries`);

  results.recentFiles = getRecentFiles();
  console.log(`  Found ${results.recentFiles.length} recent files`);

  results.storeApps = getStoreApps();
  console.log(`  Found ${results.storeApps.length} Windows Store apps`);

  results.services = getServices();
  console.log(`  Found ${results.services.length} running services`);

  results.scheduledTasks = getScheduledTasks();
  console.log(`  Found ${results.scheduledTasks.length} scheduled tasks`);

  results.networkConnections = getNetworkConnections();
  console.log(`  Found ${results.networkConnections.length} active connections`);

  // Process and categorize
  const saasUsage = extractSaasFromHistory(results.browserHistory);

  // Categorize installed apps
  const categorizedApps = results.installedPrograms.map(app => ({
    ...app,
    category: categorizeApp(app.DisplayName, app.Publisher)
  }));

  // Categorize SaaS
  const categorizedSaas = saasUsage.map(s => ({
    ...s,
    category: categorizeDomain(s.domain) || "Uncategorized"
  }));

  // Build frequency analysis from prefetch
  const appFrequency = {};
  for (const pf of results.prefetchData) {
    const name = (pf.Name || "").replace(/\.EXE$/i, "").replace(/_/g, " ");
    if (name && pf.RunCount > 0) {
      appFrequency[name] = {
        name,
        runCount: pf.RunCount || 0,
        lastRun: pf.LastRun || "",
        firstSeen: pf.Created || ""
      };
    }
  }

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`\n=== Scan complete in ${elapsed}s ===\n`);

  return {
    scanTime: new Date().toISOString(),
    scanDuration: elapsed + "s",
    system: {
      hostname: os.hostname(),
      platform: os.platform(),
      release: os.release(),
      totalMemoryGB: (os.totalmem() / (1024 ** 3)).toFixed(1),
      cpus: os.cpus()[0]?.model,
      cpuCount: os.cpus().length
    },
    installedPrograms: categorizedApps,
    runningProcesses: results.runningProcesses,
    startupPrograms: results.startupPrograms,
    saasUsage: categorizedSaas,
    browserHistory: results.browserHistory,
    prefetchData: results.prefetchData,
    appFrequency: Object.values(appFrequency).sort((a, b) => b.runCount - a.runCount),
    recentFiles: results.recentFiles,
    storeApps: results.storeApps,
    services: results.services,
    scheduledTasks: results.scheduledTasks,
    networkConnections: results.networkConnections,
    stats: {
      totalInstalled: categorizedApps.length,
      totalSaasDomains: categorizedSaas.length,
      totalProcesses: results.runningProcesses.length,
      totalStartup: results.startupPrograms.length,
      totalStoreApps: results.storeApps.length,
      totalServices: results.services.length
    }
  };
}

module.exports = { runFullScan };
