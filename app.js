/* Tri d'emails marketing : tout est traité dans le navigateur. */
(function () {
  "use strict";

  // ---------- Utilitaires ----------
  const $ = (id) => document.getElementById(id);
  const norm = (s) => String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
  const fmt = (n) => n.toLocaleString("fr-FR");
  const pct = (a, b) => (b ? Math.round((a / b) * 100) + " %" : "");
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmtDate = (d) => d.toLocaleDateString("fr-FR");

  const state = {
    fileName: "",
    headers: [],
    rows: [],          // tableaux de valeurs
    results: null,     // { items: [...], active: [...], inactive: [...] }
    tab: "kw",
    mxCache: new Map() // domaine -> "ok" | "nomx" | "nxdomain" | "unknown"
  };

  // ---------- Lecture du fichier ----------
  const drop = $("drop");
  $("file").addEventListener("change", (e) => e.target.files.length && loadFiles(e.target.files));
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => e.dataTransfer.files.length && loadFiles(e.dataTransfer.files));

  function info(el, msg, isError) {
    el.hidden = false;
    el.textContent = msg;
    el.classList.toggle("error", !!isError);
  }
  const DATA_RE = /\.(csv|txt|xlsx|xlsm|xls|ods)$/i;
  const ARCHIVE_RE = /\.(zip|rar|7z|tar|gz|tgz)$/i;

  // Bibliothèque d'archives (zip, rar, 7z) chargée seulement quand on en a besoin.
  let archiveLib = null;
  async function getArchiveLib() {
    if (!archiveLib) {
      const mod = await import(new URL("vendor/libarchive/libarchive.js", location.href).href);
      mod.Archive.init({ workerUrl: new URL("vendor/libarchive/worker-bundle.js", location.href).href });
      archiveLib = mod.Archive;
    }
    return archiveLib;
  }

  // Extrait les fichiers CSV / Excel contenus dans une archive.
  async function extractArchive(file) {
    const Archive = await getArchiveLib();
    const archive = await Archive.open(file);
    if (await archive.hasEncryptedData()) throw new Error("l'archive « " + file.name + " » est protégée par un mot de passe.");
    await archive.extractFiles();
    const entries = await archive.getFilesArray();
    const out = [];
    for (const e of entries) {
      const f = e.file;
      const full = (e.path || "") + f.name;
      if (/(^|\/)(__MACOSX|\.)/.test(full) || /^~\$/.test(f.name)) continue; // fichiers système / temporaires
      if (DATA_RE.test(f.name)) out.push({ name: full, blob: f });
    }
    try { await archive.close(); } catch (e) { /* rien */ }
    return out;
  }

  // Lit un fichier CSV / Excel et renvoie { headers, rows, sheet }.
  async function parseDataFile(name, blob) {
    const buf = await blob.arrayBuffer();
    const isCsv = /\.(csv|txt)$/i.test(name);
    let wb;
    if (isCsv) {
      let text = new TextDecoder("utf-8").decode(buf);
      if (text.includes("\uFFFD")) text = new TextDecoder("windows-1252").decode(buf); // fichiers Excel « CSV » français
      text = text.replace(/^\uFEFF/, "");
      wb = XLSX.read(text, { type: "string", raw: true });
    } else {
      wb = XLSX.read(buf, { type: "array", cellDates: true });
    }
    let aoa = [], sheet = "";
    for (const sn of wb.SheetNames) {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: true, blankrows: false });
      if (data.some((r) => r.some((v) => String(v).trim() !== ""))) { aoa = data; sheet = sn; break; }
    }
    aoa = aoa.filter((r) => r.some((v) => String(v).trim() !== ""));
    // Fichier sans ligne d'en-tête : la première ligne contient déjà un email
    const noHeader = aoa.length > 0 && aoa[0].some((v) => /\S+@\S+\.\S+/.test(String(v)));
    if (aoa.length < (noHeader ? 1 : 2)) return null;
    const width = Math.max(...aoa.map((r) => r.length));
    const seen = {};
    const headers = Array.from({ length: width }, (_, i) => {
      if (noHeader) return /\S+@\S+\.\S+/.test(String(aoa[0][i])) && !seen.Email ? (seen.Email = 1, "Email") : "Colonne " + (i + 1);
      let h = String(aoa[0][i] == null ? "" : aoa[0][i]).trim() || "Colonne " + (i + 1);
      if (seen[h]) h = h + " (" + (++seen[h]) + ")"; else seen[h] = 1;
      return h;
    });
    const rows = aoa.slice(noHeader ? 0 : 1).map((r) => Array.from({ length: width }, (_, i) => (r[i] == null ? "" : r[i])));
    return { headers, rows, sheet: !isCsv && wb.SheetNames.length > 1 ? sheet : "" };
  }

  async function loadFiles(fileList) {
    const files = Array.from(fileList);
    const fileInfo = $("file-info");
    try {
      // 1. Rassembler les fichiers de données (en ouvrant les archives)
      const sources = [];
      for (const f of files) {
        if (ARCHIVE_RE.test(f.name)) {
          info(fileInfo, "Ouverture de l'archive « " + f.name + " »…");
          const inner = await extractArchive(f);
          if (!inner.length) throw new Error("l'archive « " + f.name + " » ne contient aucun fichier CSV ou Excel.");
          sources.push(...inner);
        } else if (DATA_RE.test(f.name)) {
          sources.push({ name: f.name, blob: f });
        } else {
          throw new Error("format non pris en charge pour « " + f.name + " » (CSV, Excel, ZIP, RAR ou 7Z attendus).");
        }
      }

      // 2. Lire chaque fichier
      const parsed = [];
      for (const s of sources) {
        info(fileInfo, "Lecture de « " + s.name + " »…");
        const p = await parseDataFile(s.name, s.blob);
        if (p) parsed.push(Object.assign(p, { name: s.name }));
      }
      if (!parsed.length) throw new Error("aucune ligne de données trouvée sous l'en-tête.");

      // 3. Fusionner : colonnes réunies par nom (sans tenir compte des majuscules / accents)
      const headers = [], keyIndex = new Map();
      for (const p of parsed) for (const h of p.headers) {
        const k = norm(h);
        if (!keyIndex.has(k)) { keyIndex.set(k, headers.length); headers.push(h); }
      }
      const multi = parsed.length > 1;
      if (multi) headers.push("Fichier source");
      const rows = [];
      for (const p of parsed) {
        const map = p.headers.map((h) => keyIndex.get(norm(h)));
        for (const r of p.rows) {
          const line = new Array(headers.length).fill("");
          r.forEach((v, i) => { line[map[i]] = v; });
          if (multi) line[headers.length - 1] = p.name.split("/").pop();
          rows.push(line);
        }
      }

      state.headers = headers;
      state.rows = rows;
      state.fileName = files.length === 1 ? files[0].name.replace(/\.[^.]+$/, "") : "fusion";
      state.results = null;

      let msg;
      if (multi) {
        msg = fmt(parsed.length) + " fichiers fusionnés (" + parsed.map((p) => p.name.split("/").pop() + " : " + fmt(p.rows.length)).join(", ") +
          ") : " + fmt(rows.length) + " lignes, " + (headers.length - 1) + " colonnes. Une colonne « Fichier source » indique l'origine de chaque ligne.";
      } else {
        const p = parsed[0];
        msg = "« " + p.name + " »" + (p.sheet ? " (feuille « " + p.sheet + " »)" : "") + " : " + fmt(rows.length) + " lignes, " + headers.length + " colonnes.";
      }
      const skipped = sources.length - parsed.length;
      if (skipped) msg += " " + skipped + " fichier(s) vide(s) ignoré(s).";
      info(fileInfo, msg);
      state.original = null;
      setupParse();
      setupSettings();
    } catch (err) {
      console.error(err);
      info(fileInfo, "Impossible de lire le fichier : " + err.message, true);
    }
  }

  // ---------- Lecture : email et informations collés dans la même case ----------
  // Sépare "jean@x.fr:France" en email + informations. L'email peut être au début ou à la fin.
  function splitValue(v, sep) {
    const s = String(v == null ? "" : v).trim();
    if (!sep || !s.includes(sep)) return null;
    // Email au début : on prend le plus court début qui forme une adresse valide
    for (let k = s.indexOf(sep); k >= 0; k = s.indexOf(sep, k + 1)) {
      const email = cleanEmail(s.slice(0, k));
      if (EMAIL_RE.test(email)) return { email, parts: s.slice(k + sep.length).split(sep).map((x) => x.trim()) };
    }
    // Email à la fin : "France:jean@x.fr"
    for (let k = s.lastIndexOf(sep); k >= 0; k = s.lastIndexOf(sep, k - 1)) {
      const email = cleanEmail(s.slice(k + sep.length));
      if (EMAIL_RE.test(email)) return { email, parts: s.slice(0, k).split(sep).map((x) => x.trim()) };
      if (k === 0) break;
    }
    return null;
  }

  const SEP_CANDIDATES = [":", ";", "|", "=", "#", "/", ",", "\t", " - ", "-", " "];
  function detectGlued(col) {
    const vals = columnValues(col, 200).map(String).filter((v) => v.includes("@"));
    if (!vals.length) return null;
    if (vals.filter((v) => EMAIL_RE.test(cleanEmail(v))).length >= vals.length * 0.8) return null; // emails déjà propres
    for (const sep of SEP_CANDIDATES) {
      const ok = vals.filter((v) => { const r = splitValue(v, sep); return r && r.parts.some(Boolean); });
      if (ok.length >= vals.length * 0.5) return { sep, example: ok[0] };
    }
    return null;
  }

  function namesFromInput() {
    return $("split-names").value.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  }

  function applySplit() {
    const sep = $("split-sep").value === "\\t" ? "\t" : $("split-sep").value;
    if (!sep) { info($("parse-detect"), "Indiquez le séparateur (par exemple « : »).", true); return false; }
    if (!state.original) state.original = { headers: state.headers.slice(), rows: state.rows };
    const { headers, rows } = state.original;
    const col = +$("split-col").value;

    const splits = rows.map((r) => splitValue(r[col], sep));
    const done = splits.filter(Boolean).length;
    if (!done) {
      info($("parse-detect"), "Aucune case de la colonne « " + headers[col] + " » ne contient un email suivi de « " + sep + " ». Vérifiez le séparateur.", true);
      return false;
    }
    const maxParts = Math.max(...splits.map((s) => (s ? s.parts.length : 0)));
    let names = namesFromInput();
    if (!names.length) names = maxParts > 1 ? Array.from({ length: maxParts }, (_, i) => "Information " + (i + 1)) : ["Information"];

    const emailHeader = /mail|courriel/.test(norm(headers[col])) ? headers[col] : "Email";
    const newHeaders = headers.slice(0, col).concat([emailHeader], names, headers.slice(col + 1));
    const newRows = rows.map((r, idx) => {
      const s = splits[idx];
      let vals = s ? s.parts.slice() : [];
      // Plus de morceaux que de noms : le dernier champ reçoit le reste
      if (vals.length > names.length) vals = vals.slice(0, names.length - 1).concat([vals.slice(names.length - 1).join(sep)]);
      while (vals.length < names.length) vals.push("");
      return r.slice(0, col).concat([s ? s.email : r[col]], vals, r.slice(col + 1));
    });
    state.headers = newHeaders;
    state.rows = newRows;
    state.results = null;
    $("step-results").hidden = true;

    info($("parse-detect"), "Séparation faite sur " + fmt(done) + " ligne(s) sur " + fmt(rows.length) + " : l'email est isolé et ce qui suit « " + sep + " » va dans " +
      (names.length > 1 ? "les colonnes « " + names.join(" », « ") + " »." : "la colonne « " + names[0] + " »."));
    $("split-undo").hidden = false;
    showSplitPreview(col, names.length + 1);
    setupSettings(true);
    return true;
  }

  function showSplitPreview(col, n) {
    const box = $("split-preview");
    const cols = Array.from({ length: n }, (_, i) => col + i);
    box.querySelector("thead").innerHTML = "<tr>" + cols.map((i) => "<th>" + esc(state.headers[i]) + "</th>").join("") + "</tr>";
    box.querySelector("tbody").innerHTML = state.rows.slice(0, 5).map((r) =>
      "<tr>" + cols.map((i) => "<td>" + esc(r[i] instanceof Date ? fmtDate(r[i]) : r[i]) + "</td>").join("") + "</tr>").join("");
    box.hidden = false;
  }

  $("split-apply").addEventListener("click", applySplit);
  $("split-undo").addEventListener("click", () => {
    if (!state.original) return;
    state.headers = state.original.headers;
    state.rows = state.original.rows;
    state.original = null;
    state.results = null;
    $("step-results").hidden = true;
    $("split-undo").hidden = true;
    $("split-preview").hidden = true;
    info($("parse-detect"), "Séparation annulée : le fichier est revenu à son état d'origine.");
    setupSettings(true);
  });

  function setupParse() {
    const col = Math.max(detectEmailColumn(), 0);
    $("split-col").innerHTML = state.headers.map((h, i) => '<option value="' + i + '">' + esc(h) + "</option>").join("");
    $("split-col").value = String(col);
    $("split-undo").hidden = true;
    $("split-preview").hidden = true;
    $("split-names").value = "";
    $("understood").hidden = true;
    $("parse-detect").hidden = true;
    $("step-parse").hidden = false;

    const g = detectGlued(col);
    if (g) {
      $("split-sep").value = g.sep === "\t" ? "\\t" : g.sep;
      applySplit();
      const label = g.sep === " " ? "un espace" : g.sep === "\t" ? "une tabulation" : "« " + g.sep.trim() + " »";
      info($("parse-detect"), "Les emails sont collés à une information, séparés par " + label + " (ex. " + g.example + "). " +
        "Je les ai séparés automatiquement. Donnez un nom à ce qui suit (ex. « Pays ») dans le champ ci-dessous, puis cliquez sur « Séparer l'email et les informations ».");
    }
  }

  // ---------- Instructions écrites ou dictées ----------
  const SEP_WORDS = [
    [/deux[\s-]?points/, ":"], [/point[\s-]?virgule/, ";"], [/barre verticale|\bpipe\b/, "|"], [/barre oblique|\bslash\b/, "/"],
    [/tiret|trait d['’ ]union/, "-"], [/\begal\b|signe egal/, "="], [/diese|hashtag/, "#"], [/tabulation/, "\\t"],
    [/\bvirgule\b/, ","], [/\bespace\b/, " "]
  ];

  function understandInstructions(text) {
    const out = { notes: [] };
    let t = " " + text.replace(/\s+/g, " ") + " ";
    const n = norm(t);

    // Mot-clé
    const kw = t.match(/mots?[\s-]?cl[ée]s?\s*(?:est|sera|c['’]est|:|=)?\s*[«"“]?\s*([^.;!?\n»"”]+)/i);
    if (kw) {
      out.keyword = kw[1].replace(/\s+(?:et|puis)\s+(?:ce|la|le|les|l['’]).*$/i, "").trim();
      t = t.replace(kw[0], " ");
    }
    // Période d'inactivité
    const mo = n.match(/(\d+)\s*mois/);
    if (mo) out.months = +mo[1];
    if (/generique|noreply|no-reply/.test(n)) out.role = !/(garde|conserve|inclu)/.test(n.slice(Math.max(0, n.search(/generique|noreply/) - 30), n.search(/generique|noreply/)));

    // Séparateur
    const sym = t.match(/(?:^|\s)([:;|=#])(?:\s|$)|«\s*([^\s»])\s*»|"([^\s"])"/);
    for (const [re, sep] of SEP_WORDS) if (re.test(n)) { out.sep = sep; break; }
    if (!out.sep && sym) out.sep = sym[1] || sym[2] || sym[3];

    // Noms des informations : « c'est le pays », « s'appelle Nom », « nommer ça Ville »…
    const nameRe = /(?:c['’]est|ce sont|correspond(?:ent)?\s+(?:à|a)|s['’]appelle(?:nt)?|appel(?:le|les|lez|er|é|ée|és)|nomm(?:e|es|ez|er|é|ée|és)|intitul(?:é|ée|er|e)|représente(?:nt)?|indique(?:nt)?)\s+((?:(?:le|la|les|l['’]|du|de la|des|ça|cela|cette information|ce champ|ce qui suit|le champ|la colonne|par|comme|en)\s+)*)[«"“]?\s*([^.;!?\n»"”]+)/gi;
    let m, names = [];
    while ((m = nameRe.exec(t))) {
      const chunk = m[2].split(/\s*,\s*|\s+et\s+|\s+puis\s+|\s+ensuite\s+|\s+suivi\s+(?:de|du|de la)\s+/i)
        .map((x) => x.replace(/^(?:le|la|les|l['’]|du|de la|des|son|sa|ses)\s*/i, "").replace(/\s+(?:de la personne|du contact|de chaque.*|qui suit.*|après.*)$/i, "").trim())
        .filter((x) => x && x.split(" ").length <= 4 && !/^(deux points|point|virgule|tiret|séparateur|email|mail|l'email|le mail)$/i.test(x));
      if (chunk.length) { names = chunk; }
    }
    // Sinon : « deux points puis le pays et ensuite le nom »
    if (!names.length) {
      const nt = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      let at = -1;
      for (const [re] of SEP_WORDS) { const mm = nt.match(re); if (mm) { at = mm.index + mm[0].length; break; } }
      if (at < 0) { const mm = t.match(/[:;|=#]/); if (mm) at = mm.index + 1; }
      if (at >= 0) {
        names = t.slice(at).split(/[.!?\n]/)[0].split(/\s*,\s*|\s+et\s+|\s+puis\s+|\s+ensuite\s+|\s+suivi[es]*\s+/i)
          .map((x) => { const mm = x.match(/(?:^|\s)(?:le|la|les|l['’]|son|sa|ses)\s*([^\s].*)$/i); return mm ? mm[1].trim() : ""; })
          .filter((x) => x && x.split(" ").length <= 4 && !/^(e-?mail|mail|adresse.*)$/i.test(x));
      }
    }
    if (names.length) out.names = names.map((x) => x.charAt(0).toUpperCase() + x.slice(1));
    return out;
  }

  $("understand").addEventListener("click", () => {
    const text = $("instructions").value.trim();
    const box = $("understood");
    if (!text) { info(box, "Écrivez ou dictez d'abord vos instructions.", true); return; }
    const u = understandInstructions(text);
    const done = [];
    if (u.sep != null) { $("split-sep").value = u.sep; done.push("Séparateur après l'email : « " + (u.sep === " " ? "espace" : u.sep) + " »"); }
    if (u.names) { $("split-names").value = u.names.join(", "); done.push("Ce qui suit l'email s'appelle : " + u.names.join(", ")); }
    if (u.sep != null || u.names) {
      if (state.headers.length && applySplit()) done.push("Email et informations séparés (voir l'aperçu ci-dessous)");
    }
    if (u.keyword) { $("keyword").value = u.keyword; done.push("Mot-clé : « " + u.keyword + " »"); }
    if (u.months) {
      const opts = Array.from($("cutoff").options).map((o) => +o.value);
      const best = opts.reduce((a, b) => (Math.abs(b - u.months) < Math.abs(a - u.months) ? b : a));
      $("cutoff").value = String(best);
      done.push("Inactif sans activité depuis " + best + " mois");
    }
    if (u.role != null) { $("opt-role").checked = u.role; done.push(u.role ? "Exclure les adresses génériques" : "Garder les adresses génériques"); }

    if (!done.length) {
      info(box, "Je n'ai rien reconnu dans ces instructions. Essayez par exemple : « après l'email il y a deux points, ce qui suit c'est le pays ». Vous pouvez aussi remplir directement les champs ci-dessous.", true);
      return;
    }
    box.hidden = false;
    box.classList.remove("error");
    box.innerHTML = "<strong>Compris :</strong><ul class=\"understood-list\">" + done.map((d) => "<li>" + esc(d) + "</li>").join("") + "</ul>Vérifiez les champs ci-dessous, puis lancez le tri à l'étape 3.";
  });

  // Dictée vocale (Web Speech API)
  (function setupVoice() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const mic = $("mic"), ta = $("instructions"), note = $("mic-note");
    const label = (on) => { mic.innerHTML = on ? "■ Arrêter" : MIC_SVG + " Parler"; };
    const MIC_SVG = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true" style="vertical-align:-2px"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/></svg>';
    label(false);
    if (!SR) {
      mic.disabled = true;
      note.textContent = "La dictée vocale n'est pas disponible dans ce navigateur (utilisez Chrome, Edge ou Safari). Vous pouvez écrire vos instructions.";
      return;
    }
    let rec = null, on = false;
    mic.addEventListener("click", () => {
      if (on) { rec.stop(); return; }
      rec = new SR();
      rec.lang = "fr-FR";
      rec.continuous = true;
      rec.interimResults = true;
      const base = ta.value.trim() ? ta.value.trim() + " " : "";
      rec.onresult = (e) => {
        let txt = "";
        for (let i = 0; i < e.results.length; i++) txt += e.results[i][0].transcript;
        ta.value = base + txt.trim();
      };
      rec.onerror = (e) => {
        note.textContent = e.error === "not-allowed" || e.error === "service-not-allowed"
          ? "Le micro est bloqué. Autorisez-le dans la barre d'adresse du navigateur, puis réessayez."
          : e.error === "no-speech" ? "Je n'ai rien entendu. Cliquez sur « Parler » et réessayez." : "La dictée s'est arrêtée (" + e.error + ").";
      };
      rec.onend = () => { on = false; mic.classList.remove("on"); label(false); };
      try { rec.start(); on = true; mic.classList.add("on"); label(true); }
      catch (e) { note.textContent = "Impossible de démarrer le micro : " + e.message; }
    });
  })();

  // ---------- Détection des colonnes ----------
  function columnValues(i, limit) {
    const out = [];
    for (const r of state.rows) {
      const v = r[i];
      if (String(v).trim() !== "") out.push(v);
      if (limit && out.length >= limit) break;
    }
    return out;
  }

  function detectEmailColumn() {
    let best = -1, bestScore = 0;
    state.headers.forEach((h, i) => {
      const vals = columnValues(i, 200);
      if (!vals.length) return;
      const ratio = vals.filter((v) => String(v).includes("@")).length / vals.length;
      const score = ratio + (/mail|courriel/.test(norm(h)) ? 0.5 : 0);
      if (score > bestScore) { bestScore = score; best = i; }
    });
    return best;
  }

  function detectStatusColumn(emailCol) {
    return state.headers.findIndex((h, i) => i !== emailCol &&
      /statut|status|etat|state|abonn|subscri|actif|active|opt.?in|consent|desinscri|unsub|bounce/.test(norm(h)));
  }

  function detectCountryColumn(emailCol) {
    return state.headers.findIndex((h, i) => i !== emailCol && /pays|country|nation|^iso$|code.?pays/.test(norm(h)));
  }

  // Pays : valeur de la colonne (codes FR/BE… traduits), sinon extension du domaine.
  const capFirst = (s) => s.charAt(0).toUpperCase() + s.slice(1);
  function countryOf(row, countryCol, domain) {
    if (countryCol >= 0) {
      const raw = String(row[countryCol] == null ? "" : row[countryCol]).trim();
      if (raw) {
        const code = norm(raw);
        if (code.length === 2 && window.COUNTRY_CODES[code]) return { name: window.COUNTRY_CODES[code], guessed: false };
        return { name: capFirst(raw), guessed: false };
      }
    }
    if (domain) {
      const tld = domain.split(".").pop();
      if (window.COUNTRY_CODES[tld]) return { name: window.COUNTRY_CODES[tld], guessed: true };
    }
    return { name: "Non renseigné", guessed: false };
  }

  function detectDateColumns(emailCol) {
    const out = [];
    state.headers.forEach((h, i) => {
      if (i === emailCol) return;
      const n = norm(h);
      if (/naissance|birth|inscri|creat|signup|sign.?up|ajout|added|joined|import/.test(n)) return;
      if (!/date|dernier|derniere|last|ouvert|ouverture|open|clic|click|achat|purchase|commande|order|connexion|login|activit|visite|visit|seen|engag/.test(n)) return;
      const vals = columnValues(i, 100);
      if (!vals.length) return;
      if (vals.filter((v) => parseDate(v)).length / vals.length >= 0.5) out.push(i);
    });
    return out;
  }

  function setupSettings(keepScroll) {
    const emailCol = detectEmailColumn();
    const statusCol = detectStatusColumn(emailCol);
    const dateCols = detectDateColumns(emailCol);

    const opts = (sel, withNone) => {
      sel.innerHTML = (withNone ? '<option value="-1">— Aucune —</option>' : "") +
        state.headers.map((h, i) => '<option value="' + i + '">' + esc(h) + "</option>").join("");
    };
    opts($("col-email"), false);
    $("col-email").value = String(Math.max(emailCol, 0));
    opts($("col-status"), true);
    $("col-status").value = String(statusCol);
    opts($("col-country"), true);
    $("col-country").value = String(detectCountryColumn(emailCol));

    $("col-dates").innerHTML = state.headers.map((h, i) =>
      '<label><input type="checkbox" value="' + i + '"' + (dateCols.includes(i) ? " checked" : "") + "> " + esc(h) + "</label>").join("");
    $("col-search").innerHTML = state.headers.map((h, i) =>
      '<label><input type="checkbox" value="' + i + '" checked> ' + esc(h) + "</label>").join("");

    $("step-settings").hidden = false;
    $("step-results").hidden = true;
    if (!keepScroll) $("step-parse").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  const checked = (id) => Array.from($(id).querySelectorAll("input:checked")).map((c) => +c.value);

  // ---------- Dates ----------
  const MONTHS = { janv: 1, jan: 1, janvier: 1, fevr: 2, fev: 2, fevrier: 2, feb: 2, mars: 3, mar: 3, avr: 4, avril: 4, apr: 4, mai: 5, may: 5,
    juin: 6, jun: 6, juil: 7, juillet: 7, jul: 7, aout: 8, aug: 8, sept: 9, sep: 9, septembre: 9, oct: 10, octobre: 10,
    nov: 11, novembre: 11, dec: 12, decembre: 12 };

  function mk(y, m, d) {
    if (y < 100) y += 2000;
    const dt = new Date(y, m - 1, d);
    return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d && y > 1980 && y < 2100 ? dt : null;
  }

  function parseDate(v) {
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === "number") {
      if (v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)); // numéro de série Excel
      if (v > 1e12) return new Date(v);
      if (v > 1e9) return new Date(v * 1000);
      return null;
    }
    const s = String(v).trim();
    if (!s) return null;
    let m;
    if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return mk(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\b/))) {
      let a = +m[1], b = +m[2];
      if (a <= 12 && b > 12) [a, b] = [b, a]; // format américain mm/jj
      return mk(+m[3], b, a);
    }
    if ((m = norm(s).match(/^(\d{1,2})(?:er)?\s+([a-z]+)\.?\s+(\d{4})/))) {
      const mo = MONTHS[m[2]] || MONTHS[m[2].slice(0, 4)] || MONTHS[m[2].slice(0, 3)];
      if (mo) return mk(+m[3], mo, +m[1]);
    }
    if (/^\d{5}(\.\d+)?$/.test(s)) return parseDate(+s);
    if (/^\d{10}$|^\d{13}$/.test(s)) return parseDate(+s);
    return null;
  }

  // ---------- Contrôles d'adresse ----------
  const EMAIL_RE = /^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@([a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+(?:[a-z]{2,63}|xn--[a-z0-9-]{2,59})$/;
  const TYPOS = {
    "gmial.com": "gmail.com", "gmai.com": "gmail.com", "gmal.com": "gmail.com", "gamil.com": "gmail.com", "gnail.com": "gmail.com",
    "gmail.co": "gmail.com", "gmail.con": "gmail.com", "gmail.cm": "gmail.com", "gmaill.com": "gmail.com", "gmail.fr": "gmail.com",
    "hotmial.com": "hotmail.com", "hotmal.com": "hotmail.com", "hotmai.com": "hotmail.com", "hotmail.con": "hotmail.com", "hotmil.com": "hotmail.com",
    "hotmial.fr": "hotmail.fr", "hotmal.fr": "hotmail.fr", "hotmai.fr": "hotmail.fr",
    "yahooo.com": "yahoo.com", "yaho.com": "yahoo.com", "yahoo.con": "yahoo.com", "yahoo.fe": "yahoo.fr", "yaho.fr": "yahoo.fr",
    "outlok.com": "outlook.com", "outlook.con": "outlook.com", "outllok.com": "outlook.com", "outlok.fr": "outlook.fr",
    "orange.f": "orange.fr", "oragne.fr": "orange.fr", "wanadoo.f": "wanadoo.fr", "wanado.fr": "wanadoo.fr",
    "free.f": "free.fr", "sfr.f": "sfr.fr", "laposte.ne": "laposte.net", "icloud.co": "icloud.com", "iclod.com": "icloud.com"
  };

  const NEG_EXACT = new Set(["non", "no", "n", "false", "faux", "0", "off", "ko", "x"]);
  const NEG_SUB = ["inactif", "inactive", "desabonn", "desinscri", "unsubscrib", "unsub", "bounce", "rebond", "invalide", "invalid",
    "spam", "plainte", "complaint", "supprim", "deleted", "bloque", "blocked", "blacklist", "liste noire", "opt-out", "optout", "opt out",
    "refus", "erreur", "error", "echec", "failed", "cleaned", "archiv", "ferme", "closed", "desactiv", "disabled", "non abonne", "pas abonne"];

  function negativeStatus(v) {
    const n = norm(v);
    if (!n) return false;
    if (NEG_EXACT.has(n)) return true;
    return NEG_SUB.some((t) => n.includes(t));
  }

  function cleanEmail(v) {
    return String(v == null ? "" : v).trim().toLowerCase().replace(/^mailto:/, "").replace(/^[<"'\s]+|[>"'\s]+$/g, "");
  }

  // ---------- Vérification DNS (MX) ----------
  async function dohQuery(domain, type) {
    const urls = [
      "https://cloudflare-dns.com/dns-query?name=" + encodeURIComponent(domain) + "&type=" + type,
      "https://dns.google/resolve?name=" + encodeURIComponent(domain) + "&type=" + type
    ];
    for (const url of urls) {
      try {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), 8000);
        const res = await fetch(url, { headers: { accept: "application/dns-json" }, signal: ctrl.signal });
        clearTimeout(t);
        if (res.ok) return await res.json();
      } catch (e) { /* on essaie le service suivant */ }
    }
    return null;
  }

  async function checkDomain(domain) {
    if (state.mxCache.has(domain)) return state.mxCache.get(domain);
    let result = "unknown";
    const mx = await dohQuery(domain, "MX");
    if (mx) {
      if (mx.Status === 3) result = "nxdomain";
      else if (mx.Status === 0) {
        const recs = (mx.Answer || []).filter((a) => a.type === 15);
        if (recs.length) {
          // « 0 . » = domaine qui déclare explicitement ne pas recevoir d'emails
          result = recs.every((r) => /^\d+\s+\.?$/.test(String(r.data).trim())) ? "nomx" : "ok";
        } else {
          const a = await dohQuery(domain, "A"); // repli autorisé par la norme : pas de MX mais une adresse IP
          result = a && a.Status === 0 && (a.Answer || []).some((x) => x.type === 1) ? "ok" : (a ? "nomx" : "unknown");
        }
      }
    }
    if (result !== "unknown") state.mxCache.set(domain, result);
    return result;
  }

  async function checkDomains(domains, onProgress) {
    const list = Array.from(domains);
    let done = 0, idx = 0;
    const out = new Map();
    async function worker() {
      while (idx < list.length) {
        const d = list[idx++];
        out.set(d, await checkDomain(d));
        onProgress(++done, list.length);
      }
    }
    await Promise.all(Array.from({ length: Math.min(10, list.length) }, worker));
    return out;
  }

  // ---------- Tri ----------
  $("run").addEventListener("click", run);

  async function run() {
    const btn = $("run"), progress = $("progress");
    btn.disabled = true;
    try {
      const emailCol = +$("col-email").value;
      const statusCol = +$("col-status").value;
      const dateCols = checked("col-dates");
      const months = +$("cutoff").value;
      const noDateDrop = $("nodate").value === "drop";
      const optMx = $("opt-mx").checked, optDisp = $("opt-disposable").checked, optRole = $("opt-role").checked;

      const cutoff = new Date();
      cutoff.setMonth(cutoff.getMonth() - months);

      const countryCol = +$("col-country").value;
      const countryNames = new Map(); // clé normalisée -> nom affiché (première forme rencontrée)
      const seen = new Set();
      const items = state.rows.map((row, idx) => {
        const email = cleanEmail(row[emailCol]);
        const it = { idx, row, email, reason: null, cat: null, last: null };
        const at = email.lastIndexOf("@");
        const c = countryOf(row, countryCol, at > 0 ? email.slice(at + 1) : "");
        const key = norm(c.name);
        if (!countryNames.has(key)) countryNames.set(key, c.name);
        it.country = countryNames.get(key);
        it.countryGuessed = c.guessed;
        const fail = (cat, reason) => { it.cat = cat; it.reason = reason || cat; };

        if (!email) { fail("Email manquant"); return it; }
        if (!EMAIL_RE.test(email)) {
          if (/^[^\s@]+@[^\s@]+\.[a-z]{2,}[:;|=#\/,]/i.test(email)) fail("Information collée à l'email", "Information collée à l'email : séparez-la à l'étape 2");
          else fail("Adresse mal formée");
          return it;
        }
        if (seen.has(email)) { fail("Doublon"); return it; }
        seen.add(email);

        const [local, domain] = [email.slice(0, email.lastIndexOf("@")), email.slice(email.lastIndexOf("@") + 1)];
        it.domain = domain;
        if (TYPOS[domain]) { fail("Faute de frappe probable", "Faute de frappe probable (" + domain + " → " + TYPOS[domain] + ")"); return it; }
        if (optDisp && window.DISPOSABLE_DOMAINS.has(domain)) { fail("Adresse jetable"); return it; }
        if (optRole && window.ROLE_PREFIXES.has(local.split("+")[0])) { fail("Adresse générique"); return it; }
        if (statusCol >= 0 && negativeStatus(row[statusCol])) { fail("Statut négatif", "Statut : " + String(row[statusCol]).trim()); return it; }

        if (dateCols.length) {
          for (const c of dateCols) {
            const d = parseDate(row[c]);
            if (d && (!it.last || d > it.last)) it.last = d;
          }
          if (!it.last && noDateDrop) { fail("Aucune date d'activité"); return it; }
          if (it.last && it.last < cutoff) {
            fail("Inactif depuis plus de " + months + " mois", "Aucune activité depuis plus de " + months + " mois (dernière : " + fmtDate(it.last) + ")");
            return it;
          }
        }
        return it;
      });

      let unknownDomains = 0;
      if (optMx) {
        const domains = new Set(items.filter((i) => !i.cat).map((i) => i.domain));
        info(progress, "Vérification des domaines : 0 / " + fmt(domains.size));
        const res = await checkDomains(domains, (d, t) => { progress.textContent = "Vérification des domaines : " + fmt(d) + " / " + fmt(t); });
        for (const it of items) {
          if (it.cat) continue;
          const r = res.get(it.domain);
          if (r === "nxdomain") { it.cat = "Domaine inexistant"; it.reason = "Domaine inexistant (" + it.domain + ")"; }
          else if (r === "nomx") { it.cat = "Domaine sans serveur mail"; it.reason = "Le domaine " + it.domain + " ne reçoit pas d'emails"; }
        }
        unknownDomains = Array.from(res.values()).filter((v) => v === "unknown").length;
      }

      state.results = {
        items,
        active: items.filter((i) => !i.cat),
        inactive: items.filter((i) => i.cat),
        unknownDomains,
        hasDates: dateCols.length > 0,
        countryCol,
        guessedCountries: items.some((i) => i.countryGuessed)
      };
      state.country = "";
      progress.hidden = true;
      $("step-results").hidden = false;
      render();
      $("step-results").scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (err) {
      console.error(err);
      info(progress, "Erreur pendant le tri : " + err.message, true);
    } finally {
      btn.disabled = false;
    }
  }

  // ---------- Mot-clé ----------
  function keywordTerms() {
    return $("keyword").value.split(/[,;]/).map(norm).filter(Boolean);
  }

  function matchKeyword(it, terms, cols) {
    const hits = [];
    for (const c of cols) {
      let v = it.row[c];
      if (v instanceof Date) v = fmtDate(v);
      const n = norm(v);
      if (n && terms.some((t) => n.includes(t))) hits.push(state.headers[c]);
    }
    return hits;
  }

  let kwTimer;
  $("keyword").addEventListener("input", () => { clearTimeout(kwTimer); kwTimer = setTimeout(() => state.results && render(), 200); });
  $("col-search").addEventListener("change", () => state.results && render());
  $("keyword").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); state.results ? render() : run(); } });

  // ---------- Affichage ----------
  function currentLists() {
    const r = state.results;
    const terms = keywordTerms();
    const cols = checked("col-search");
    let kw = r.active;
    if (terms.length) {
      kw = [];
      for (const it of r.active) {
        const hits = matchKeyword(it, terms, cols);
        if (hits.length) kw.push(Object.assign({}, it, { hits }));
      }
    }
    return { terms, kw, active: r.active, inactive: r.inactive };
  }

  // Liste de l'onglet affiché, filtrée par pays et triée.
  function shownList(L) {
    let list = L[state.tab];
    if (state.country) list = list.filter((i) => i.country === state.country);
    const sort = $("sort").value;
    const unknownLast = (c) => (c === "Non renseigné" ? "￿" : c);
    if (sort === "country") list = list.slice().sort((a, b) => unknownLast(a.country).localeCompare(unknownLast(b.country), "fr") || a.idx - b.idx);
    else if (sort === "email") list = list.slice().sort((a, b) => a.email.localeCompare(b.email, "fr"));
    return list;
  }

  function render() {
    const r = state.results;
    const L = currentLists();
    const total = r.items.length;
    const kwLabel = L.terms.length ? "« " + $("keyword").value.trim() + " »" : "";

    $("s-total").textContent = fmt(total);
    $("s-active").textContent = fmt(r.active.length);
    $("s-active-p").textContent = pct(r.active.length, total) + " du fichier";
    $("s-inactive").textContent = fmt(r.inactive.length);
    $("s-inactive-p").textContent = pct(r.inactive.length, total) + " du fichier";
    $("s-kw-label").textContent = L.terms.length ? "Actifs liés à " + kwLabel : "Actifs liés au mot-clé";
    $("s-kw").textContent = L.terms.length ? fmt(L.kw.length) : "—";
    $("s-kw-p").textContent = L.terms.length ? pct(L.kw.length, r.active.length) + " des actifs" : "Saisissez un mot-clé";
    document.querySelectorAll("button.stat").forEach((b) => b.classList.toggle("selected", b.dataset.tab === state.tab));

    let sum = "Sur les <strong>" + fmt(total) + "</strong> contacts du fichier, <strong>" + fmt(r.active.length) + "</strong> sont actifs";
    sum += r.active.length ? " (" + pct(r.active.length, total) + ")." : ".";
    if (L.terms.length) sum += " Parmi les actifs, <strong>" + fmt(L.kw.length) + "</strong> correspondent au mot-clé " + esc(kwLabel) + ".";
    if (!r.hasDates) sum += "<br><small>Aucune colonne de date d'activité n'est utilisée : le tri repose sur la validité des adresses et le statut.</small>";
    if (r.unknownDomains) sum += "<br><small>" + fmt(r.unknownDomains) + " domaine(s) n'ont pas pu être vérifiés (connexion) : leurs contacts sont gardés comme actifs.</small>";
    if (r.guessedCountries) sum += "<br><small>" + (r.countryCol >= 0 ? "Quand la colonne pays est vide, le" : "Le") + " pays est déduit de l'extension du domaine (.fr, .be…). Les adresses en .com, .net… restent « Non renseigné ».</small>";
    $("summary").innerHTML = sum;

    const counts = {};
    r.inactive.forEach((i) => { counts[i.cat] = (counts[i.cat] || 0) + 1; });
    $("reasons").innerHTML = Object.entries(counts).sort((a, b) => b[1] - a[1])
      .map(([k, v]) => "<span>" + esc(k) + " : <strong>" + fmt(v) + "</strong></span>").join("");

    const tabNames = { kw: "Actifs + mot-clé", active: "Tous les actifs", inactive: "Inactifs / exclus" };
    document.querySelectorAll(".tab").forEach((t) => {
      const n = t.dataset.tab;
      const count = n === "kw" ? (L.terms.length ? L.kw.length : null) : L[n].length;
      t.textContent = tabNames[n] + (count == null ? "" : " (" + fmt(count) + ")");
      t.classList.toggle("active", n === state.tab);
    });

    // Répartition par pays de la liste de l'onglet affiché
    const byCountry = new Map();
    L[state.tab].forEach((i) => byCountry.set(i.country, (byCountry.get(i.country) || 0) + 1));
    const entries = Array.from(byCountry.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "fr"));
    if (state.country && !byCountry.has(state.country)) state.country = "";
    $("countries-title").textContent = "Répartition par pays : " + tabNames[state.tab].toLowerCase();
    $("countries").innerHTML = entries.length
      ? '<button data-c=""' + (state.country ? "" : ' class="selected"') + ">Tous : " + fmt(L[state.tab].length) + "</button>" +
        entries.map(([c, n]) => '<button data-c="' + esc(c) + '"' + (c === state.country ? ' class="selected"' : "") + ">" + esc(c) + " : <strong>" + fmt(n) + "</strong></button>").join("")
      : "<span>Aucun contact</span>";
    const sel = $("country-filter");
    sel.innerHTML = '<option value="">Tous les pays</option>' + entries.map(([c, n]) => '<option value="' + esc(c) + '">' + esc(c) + " (" + fmt(n) + ")</option>").join("");
    sel.value = state.country;

    renderTable(L);
  }

  const setTab = (t) => { state.tab = t; render(); };
  document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => setTab(t.dataset.tab)));
  document.querySelectorAll("button.stat").forEach((b) => b.addEventListener("click", () => {
    setTab(b.dataset.tab);
    $("table").scrollIntoView({ behavior: "smooth", block: "nearest" });
  }));
  $("countries").addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (b) { state.country = b.dataset.c; render(); }
  });
  $("country-filter").addEventListener("change", (e) => { state.country = e.target.value; render(); });
  $("sort").addEventListener("change", () => state.results && render());

  function tableData(L) {
    const emailCol = +$("col-email").value;
    const countryCol = state.results.countryCol;
    const tab = state.tab;
    const list = shownList(L);
    const others = state.headers.map((h, i) => i).filter((i) => i !== emailCol && i !== countryCol);
    const head = ["Email", "Pays"];
    if (tab === "inactive") head.push("Raison");
    if (tab === "kw" && L.terms.length) head.push("Mot-clé trouvé dans");
    if (state.results.hasDates && tab !== "inactive") head.push("Dernière activité");
    others.forEach((i) => head.push(state.headers[i]));

    const body = list.map((it) => {
      const line = [it.email || String(it.row[emailCol]), it.country];
      if (tab === "inactive") line.push(it.reason);
      if (tab === "kw" && L.terms.length) line.push((it.hits || []).join(", "));
      if (state.results.hasDates && tab !== "inactive") line.push(it.last ? fmtDate(it.last) : "");
      others.forEach((i) => { const v = it.row[i]; line.push(v instanceof Date ? fmtDate(v) : v); });
      return line;
    });
    return { head, body };
  }

  const MAX_ROWS = 1000;
  function renderTable(L) {
    const { head, body } = tableData(L);
    const table = $("table");
    table.querySelector("thead").innerHTML = "<tr>" + head.map((h) => "<th>" + esc(h) + "</th>").join("") + "</tr>";
    const reasonIdx = state.tab === "inactive" ? 2 : -1;
    const grouped = $("sort").value === "country";
    const groupCounts = {};
    if (grouped) body.forEach((r) => { groupCounts[r[1]] = (groupCounts[r[1]] || 0) + 1; });
    let html = "", prev = null;
    for (const r of (state.showAll ? body : body.slice(0, MAX_ROWS))) {
      if (grouped && r[1] !== prev) {
        prev = r[1];
        html += '<tr class="group"><td colspan="' + head.length + '">' + esc(r[1]) + " · " + fmt(groupCounts[r[1]]) + " contact(s)</td></tr>";
      }
      html += "<tr>" + r.map((v, i) => "<td" + (i === reasonIdx ? ' class="reason"' : "") + ">" + esc(v) + "</td>").join("") + "</tr>";
    }
    table.querySelector("tbody").innerHTML = html;
    const note = $("table-note");
    if (!body.length) note.textContent = "Aucun contact dans cette liste.";
    else if (body.length > MAX_ROWS && !state.showAll) {
      note.innerHTML = "Aperçu des " + fmt(MAX_ROWS) + " premières lignes sur " + fmt(body.length) + ". L'export contient toutes les lignes. ";
      const more = document.createElement("button");
      more.className = "btn ghost"; more.textContent = "Afficher les " + fmt(body.length) + " lignes";
      more.onclick = () => { state.showAll = true; render(); };
      note.appendChild(more);
    }
    else if (state.tab === "kw" && !L.terms.length) note.textContent = "Aucun mot-clé saisi : tous les actifs sont affichés.";
    else note.textContent = "";
  }

  // ---------- Export ----------
  function exportName(ext) {
    const terms = keywordTerms();
    const slug = (s) => norm(s).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40);
    const part = { kw: "actifs" + (terms.length ? "-" + slug(terms.join("-")) : ""), active: "actifs", inactive: "inactifs" }[state.tab];
    return slug(state.fileName) + "-" + part + (state.country ? "-" + slug(state.country) : "") + "." + ext;
  }

  $("export-csv").addEventListener("click", () => {
    const { head, body } = tableData(currentLists());
    const cell = (v) => { const s = String(v == null ? "" : v); return /[;"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const csv = "\uFEFF" + [head, ...body].map((r) => r.map(cell).join(";")).join("\r\n");
    download(new Blob([csv], { type: "text/csv;charset=utf-8" }), exportName("csv"));
  });

  $("export-xlsx").addEventListener("click", () => {
    const { head, body } = tableData(currentLists());
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ...body]), "Résultat");
    XLSX.writeFile(wb, exportName("xlsx"));
  });

  $("copy-emails").addEventListener("click", async () => {
    const { body } = tableData(currentLists());
    const text = body.map((r) => r[0]).join("\n");
    const btn = $("copy-emails");
    try { await navigator.clipboard.writeText(text); }
    catch (e) {
      const ta = document.createElement("textarea");
      ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); ta.remove();
    }
    btn.textContent = fmt(body.length) + " emails copiés";
    setTimeout(() => { btn.textContent = "Copier les emails"; }, 2000);
  });

  function download(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
})();
