/* SimplyPOS — Εργαλείο Παραγγελιών (μόνο ανάγνωση· δεν καταχωρεί τίποτα στο SimplyPOS)
   Έκδοση 18/10/2026: Excel με τελικούς αριθμούς (χωρίς υπολογισμούς), γκρι στήλη ποσότητας,
   χωρίς είδη με ποσότητα 0, κίτρινο κελί «ΔΙΑΓΡΑΦΗ γραμμών με ποσότητα έως» στη Σύνοψη,
   επιλογή «Αφαίρεση τρέχοντος αποθέματος» (ΟΧΙ = ποσότητα = πωλήσεις περιόδου),
   επιλογή «Κατασκευαστής» (τα είδη του κατασκευαστή παίρνονται από τη λίστα ειδών /Product/ProductListInfinite,
   γιατί το φίλτρο κατασκευαστή της αναφοράς πωλήσεων του SimplyPOS φέρνει λάθος είδη) + στήλη Κατασκευαστής,
   επιλογή «Εκτύπωση» (ΣΥΝΟΨΗ / ΣΥΝΟΛΙΚΗ / ΠΡΟΜΗΘΕΥΤΕΣ): A4 όρθιο, μόνο Κωδ. προμηθευτή, Είδος, Ποσότητα, Σύνολο με ΦΠΑ,
   επιλογή «Στρογγυλοποίηση κιβωτίων» (πεδίο «Ποσότητα» στις Πληροφορίες του είδους· χωρίς τιμή = ως έχει),
   επιλογή «Καταχώρηση παραγγελίας» → ΡΑΓΚΟΥΣΗΣ / Mmmgroup (site): ο ίδιος σελιδοδείκτης στο site του προμηθευτή
   βάζει την παραγγελία στο ΚΑΛΑΘΙ (ποτέ δεν τη στέλνει — την αποστολή την κάνει ο χρήστης).
   ΖΟΥΜΗΣ / ΜΙΧΑΛΙΟΣ / ΜΙΕΤΣΕΛ / ΜΑΡΙΝΟΠΟΥΛΟΣ ΙΚΕ (28/10/2026): e-mail από rrms2021@gmail.com προς cavazoumis@gmail.com /
   mixalioscustomers@gmail.com / paros.mietsel@gmail.com / marinopl@yahoo.com. ΣΤΕΦΑΝΙΔΗΣ Ι Α.Ε.: info@stefanidisgroup.gr. ASTERAS: ktimaroussoswinery@gmail.com. Με ρυθμισμένη «Αυτόματη αποστολή»
   (Google Apps Script του χρήστη, σύνδεσμος/κλειδί μόνο στον browser) στέλνεται με ένα πάτημα· αλλιώς έτοιμο μήνυμα στο Gmail. */
(async () => {
  'use strict';
  const PID = 'sp-orders-panel';
  // Προμηθευτές με site: η παραγγελία πάει στο καλάθι του site (ποτέ δεν στέλνεται αυτόματα)
  const SEND = {
    ragousis: { vid: '13473', label: 'ΡΑΓΚΟΥΣΗΣ', url: 'https://ragousiscava.it-support.gr/cart' },
    mmm: { vid: '9941', label: 'Mmmgroup', url: 'https://www.mmmgroup.gr/' },
    mas: { vid: '13164', label: 'ΜΑΣΟΥΤΗΣ', url: 'https://expressapp.masoutis.gr/home' },
    // Προμηθευτής με e-mail: ανοίγει έτοιμο μήνυμα στο Gmail· την αποστολή την πατάει ο χρήστης
    zoumis: { vid: '10198', label: 'ΖΟΥΜΗΣ', email: 'cavazoumis@gmail.com', from: 'rrms2021@gmail.com' },
    michalios: { vid: '10194', label: 'ΜΙΧΑΛΙΟΣ', email: 'mixalioscustomers@gmail.com', from: 'rrms2021@gmail.com' },
    mietsel: { vid: '9982', label: 'ΜΙΕΤΣΕΛ', email: 'paros.mietsel@gmail.com', from: 'rrms2021@gmail.com' },
    marinopoulos: { vid: '19775', label: 'ΜΑΡΙΝΟΠΟΥΛΟΣ', email: 'marinopl@yahoo.com', from: 'rrms2021@gmail.com' },
    stefanidis: { vid: '21663', label: 'ΣΤΕΦΑΝΙΔΗΣ', email: 'info@stefanidisgroup.gr', from: 'rrms2021@gmail.com' },
    asteras: { vid: '10187', label: 'ASTERAS', email: 'ktimaroussoswinery@gmail.com', from: 'rrms2021@gmail.com' }
  };
  if (/ragousiscava\./i.test(location.hostname)) { await supplierModule('ragousis'); return; }
  if (/(^|\.)mmmgroup\.gr$/i.test(location.hostname)) { await supplierModule('mmm'); return; }
  if (/^expressapp\.masoutis\.gr$/i.test(location.hostname)) { await supplierModule('mas'); return; }
  if (!/simplypos\.com$/i.test(location.hostname)) {
    alert('Ανοίξτε πρώτα το app.simplypos.com (ή το site του προμηθευτή) συνδεδεμένοι και πατήστε ξανά τον σελιδοδείκτη «Παραγγελίες».');
    return;
  }
  const existing = document.getElementById(PID);
  if (existing) { existing.style.display = 'flex'; return; }

  // ---------- helpers ----------
  const $id = (x) => document.getElementById(x);
  const pad = (n) => String(n).padStart(2, '0');
  const isoToGr = (iso) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
  const isoShiftYear = (iso, k) => {
    let [y, m, d] = iso.split('-').map(Number); y += k;
    const last = new Date(y, m, 0).getDate(); if (d > last) d = last;
    return `${y}-${pad(m)}-${pad(d)}`;
  };
  const isoOf = (dt) => `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
  const fmtEur = (v) => (v || 0).toLocaleString('el-GR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
  const fmtInt = (v) => Math.round(v || 0).toLocaleString('el-GR');
  const r2 = (v) => Math.round((v + Number.EPSILON) * 100) / 100;
  const r4 = (v) => Math.round((v + Number.EPSILON) * 10000) / 10000;
  const normName = (s) => String(s ?? '').replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const store = { get(k, d) { try { const v = localStorage.getItem('spOrders.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
                  set(k, v) { try { localStorage.setItem('spOrders.' + k, JSON.stringify(v)); } catch (e) {} } };

  function param(obj, prefix) { // jQuery-style serializer
    const out = [];
    const add = (k, v) => out.push(encodeURIComponent(k) + '=' + encodeURIComponent(v == null ? '' : v));
    const walk = (v, k) => {
      if (Array.isArray(v)) v.forEach((x, i) => walk(x, `${k}[${typeof x === 'object' && x !== null ? i : ''}]`));
      else if (v !== null && typeof v === 'object') Object.keys(v).forEach((kk) => walk(v[kk], k ? `${k}[${kk}]` : kk));
      else add(k, v);
    };
    walk(obj, prefix || '');
    return out.join('&').replace(/%20/g, '+');
  }
  const loadScript = (src) => new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Δεν φορτώθηκε: ' + src));
    document.head.appendChild(s);
  });

  // ---------- SimplyPOS data ----------
  async function getHtmlDoc(url) {
    const r = await fetch(url, { credentials: 'same-origin' });
    if (!r.ok) throw new Error('Σφάλμα ' + r.status + ' στο ' + url);
    return new DOMParser().parseFromString(await r.text(), 'text/html');
  }
  async function loadMeta() {
    const [rep, po] = await Promise.all([getHtmlDoc('/Report/SalesByProduct'), getHtmlDoc('/PurchaseOrderBiz/Create')]);
    const stores = [...rep.querySelectorAll('#CompanyIds option')].map((o) => ({ id: o.value, name: o.textContent.trim() })).filter((s) => s.id);
    const vendors = {};
    [...po.querySelectorAll('#VendorId option')].forEach((o) => { if (o.value) vendors[o.value] = o.textContent.trim(); });
    // Κατασκευαστές: id + όνομα. Ένα είδος μπορεί να ανήκει σε πολλούς κατασκευαστές, γι' αυτό
    // το φίλτρο γίνεται από τον server του SimplyPOS (ProductBrandId) και όχι από τη στήλη της αναφοράς.
    const brands = [...rep.querySelectorAll('#ProductBrandId option')]
      .map((o) => ({ id: o.value, name: normName(o.textContent) })).filter((b) => b.id && b.name && b.name !== 'Όλοι');
    const activeName = (document.querySelector('span.selectstoremobile')?.textContent ||
                        rep.querySelector('span.selectstoremobile')?.textContent || '').trim();
    const active = stores.find((s) => s.name === activeName) || stores[0];
    const brandName = {}; brands.forEach((b) => { brandName[b.id] = b.name; });
    return { stores, vendors, brands, brandName, active };
  }
  // Σημ.: draw=1 επιστρέφει πάντα κενό από τον server (αρχική φόρτωση DataTables), γι αυτό draw=2.
  async function salesByCode(fromIso, toIso, companyId, brandId) {
    const columns = [];
    for (let i = 0; i < 14; i++) columns.push({ data: i, name: '', searchable: true, orderable: i !== 0, search: { value: '', regex: false } });
    const body = {
      draw: 2, columns, order: [{ column: 1, dir: 'asc' }], start: 0, length: 100000, search: { value: '', regex: false },
      action: 'filter', CompanyId: companyId, ProductId: '', order_date_from: isoToGr(fromIso), order_date_to: isoToGr(toIso),
      SubCategoryId: '', ProductBrandId: brandId || '', VendorId: '', TotalSalesFrom: '', TotalSalesTo: '', QtyFrom: '', QtyTo: '',
      CostOfGoodsFrom: '', CostOfGoodsTo: '', NetSalesFrom: '', NetSalesTo: '', ProfitFrom: '', ProfitTo: '',
      ProfitMarginFrom: '', ProfitMarginTo: '', VATFrom: '', VATTo: ''
    };
    const r = await fetch('/api/ApiReports/GetSaleByProductReport', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
      body: param(body)
    });
    if (!r.ok) throw new Error('Σφάλμα αναφοράς πωλήσεων (' + r.status + ')');
    const j = await r.json();
    const m = {}, b = {};
    (j.data || []).forEach((row) => {
      const cd = String(row[1]).trim(); m[cd] = (m[cd] || 0) + (parseFloat(row[7]) || 0);
      const br = normName(row[4]); if (br && !b[cd]) b[cd] = br;   // στήλη «Όνομα κατασκευαστή» της αναφοράς
    });
    return { q: m, b };
  }
  // Είδη ενός κατασκευαστή από τη ΛΙΣΤΑ ΕΙΔΩΝ (Αποθήκη › Είδη), σελίδα-σελίδα.
  // (Το φίλτρο κατασκευαστή της αναφοράς πωλήσεων του SimplyPOS φέρνει λάθος είδη, γι' αυτό δεν χρησιμοποιείται.)
  async function brandProductIds(brandId) {
    const ids = new Set();
    for (let p = 1; p < 500; p++) {
      const qs = new URLSearchParams({ page: p, nameCode: '', Msg: 'null', SubCategoryId: '', ProductTypeId: '', ProductBrandId: brandId,
        ProductBrand_ID: '', VendorId: '', StockLevelId: '', Category_ID: '', FeaturedStatusId: '', SortDirection: '', SortBy: '', TabStatus: 1 });
      const r = await fetch('/Product/ProductListInfinite?' + qs.toString(), { credentials: 'same-origin' });
      if (!r.ok) throw new Error('Σφάλμα λίστας ειδών κατασκευαστή (' + r.status + ')');
      const found = [...(await r.text()).matchAll(/Product_ID=([0-9a-f-]{36})/gi)].map((m) => m[1].toLowerCase());
      const before = ids.size; found.forEach((x) => ids.add(x));
      if (!found.length || ids.size === before) break;
    }
    return ids;
  }
  // «Ποσότητα» κιβωτίου από τη σελίδα Πληροφορίες του είδους (Product/Edit, πεδίο PackageQuantity)
  async function packQuantities(pids, onProgress) {
    const out = {}; let i = 0, done = 0;
    const worker = async () => {
      while (i < pids.length) {
        const pid = pids[i++];
        try {
          const h = await fetch('/Product/Edit?Product_ID=' + encodeURIComponent(pid), { credentials: 'same-origin' }).then((r) => r.text());
          const m = h.match(/id="PackageQuantity"[^>]*value="([^"]*)"/) || h.match(/value="([^"]*)"[^>]*id="PackageQuantity"/);
          const v = m ? parseFloat(String(m[1]).replace(',', '.')) : NaN;
          if (v > 0) out[String(pid).toLowerCase()] = v;
        } catch (e) { /* αγνοείται — το είδος μένει ως έχει */ }
        done++; if (onProgress && done % 20 === 0) onProgress(done, pids.length);
      }
    };
    await Promise.all(Array.from({ length: 10 }, worker));
    return out;
  }
  async function purchaseProducts(storeId) {
    const r = await fetch('/api/ApiProducts/GetPurchaseProducts?CompId=' + encodeURIComponent(storeId), { credentials: 'same-origin' });
    if (!r.ok) throw new Error('Σφάλμα λίστας ειδών (' + r.status + ')');
    return r.json();
  }

  // ---------- calculation (γίνεται εδώ· στο Excel γράφονται μόνο τελικοί αριθμοί) ----------
  // sYs = πωλήσεις ίδιας περιόδου στα προηγούμενα έτη [πέρσι, πρόπερσι, …] (κενό αν δεν ζητήθηκε Μ.Ο.)
  function compute(o, meta, P, sCur, sYs, brandOf, packOf) {
    sYs = sYs || [];
    brandOf = brandOf || {}; packOf = packOf || {};
    const vname = (v) => (v == null || v === '' ? '' : (meta.vendors[v] || ('Προμηθευτής #' + v)));
    const codes = new Set([...Object.keys(sCur), ...sYs.flatMap((m) => Object.keys(m))]);
    const byCd = {}; P.forEach((x) => { byCd[String(x.cd)] = x; });
    const rows = [], skippedNoCost = [], unknown = [];
    for (const cd of codes) {
      const x = byCd[cd];
      if (!x) { if ((sCur[cd] || 0) > 0) unknown.push(cd); continue; }
      const br = brandOf[cd] || '';   // (το φίλτρο κατασκευαστή έχει ήδη γίνει στις αναφορές πωλήσεων)
      const cur = sCur[cd] || 0;
      // Μ.Ο. μόνο των ετών που το είδος είχε πωλήσεις στην ίδια περίοδο (π.χ. 3 από τα 5 → διαιρεί με 3)
      const yv = sYs.map((m) => m[cd] || 0), yN = yv.filter((v) => v > 0).length;
      const yAvg = yN ? yv.reduce((a, v) => a + v, 0) / yN : 0;
      const base = o.avg2y ? yAvg : cur;
      if (!(base > o.minSales)) continue;
      const cstk = Number(x.cstk) || 0, dstk = Number(x.dstk) || 0;
      const qRaw = Math.max(0, Math.ceil((o.subStock ? base - cstk : base) - 1e-9));
      let q = qRaw; const notes = [];
      if (o.avg2y && yN) notes.push('Μ.Ο. ' + yN + (yN === 1 ? ' έτους' : ' ετών') + ' (' + yv.map((v, i) => (o.yearsLbl ? o.yearsLbl[i] : 'έτος -' + (i + 1)) + ': ' + Math.round(v * 100) / 100).join(', ') + ').');
      if (!o.overDesired) {
        if (dstk > 0) q = Math.min(qRaw, Math.max(0, Math.ceil(dstk - cstk - 1e-9)));
        else notes.push('Χωρίς επιθυμητό απόθεμα — δεν εφαρμόστηκε όριο.');
      }
      if (q <= 0) continue; // είδη με ποσότητα 0 δεν μπαίνουν στο Excel
      // στρογγυλοποίηση στο πλησιέστερο κιβώτιο (τουλάχιστον 1)· χωρίς «Ποσότητα» στο είδος → ως έχει
      const qPre = q, pk = o.boxRound ? (packOf[String(x.pid || '').toLowerCase()] || 0) : 0; let boxes = null;
      if (pk > 0) {
        boxes = Math.max(1, Math.round(q / pk)); q = r4(boxes * pk);
        if (q !== qPre) notes.push(`Στρογγυλοποίηση: ${qPre} τεμ. → ${boxes} κιβ. × ${pk} τεμ.`);
      }
      // offers per vendor
      const offers = {}; let suspicious = false;
      (x.sCdList || []).forEach((s) => {
        const cost = Number(s.cost) || 0, disc = Number(s.disc) || 0, net = cost * (1 - disc / 100);
        if (cost <= 0) return;
        if (disc >= 90 || net < 0.02) { suspicious = true; return; }
        const k = String(s.vid);
        if (!offers[k] || net < offers[k].net) offers[k] = { vid: s.vid, cost, disc, net, code: s.code || '' };
      });
      const mainVid = x.vid != null ? String(x.vid) : '';
      const mainCost = Number(x.pr) || 0;
      if (mainVid && !offers[mainVid] && mainCost > 0.02) offers[mainVid] = { vid: x.vid, cost: mainCost, disc: 0, net: mainCost, code: '' };
      const list = Object.values(offers).sort((a, b) => a.net - b.net);
      let pick;
      if (o.cheapest) pick = list[0];
      else if (o.vendor) {
        // ΟΧΙ + επιλεγμένος προμηθευτής → μόνο τα είδη που έχουν αυτόν τον προμηθευτή
        // δηλωμένο στην καρτέλα «Προμηθευτής» του είδους στο SimplyPOS (κύριο ή δευτερεύοντα)
        const want = String(o.vendor);
        const decl = (x.sCdList || []).filter((s) => String(s.vid) === want);
        if (mainVid !== want && !decl.length) continue;
        pick = offers[want];
        if (!pick) {
          const fb = (offers[mainVid] && offers[mainVid].net > 0) ? offers[mainVid] : (mainCost > 0.02 ? { vid: x.vid, cost: mainCost, disc: 0, net: mainCost } : list[0]);
          if (fb) {
            pick = { vid: want, cost: fb.cost, disc: fb.disc || 0, net: fb.net, code: (decl[0] && decl[0].code) || '' };
            notes.push('Ο προμηθευτής είναι δηλωμένος χωρίς τιμή στο SimplyPOS — κόστος από ' + (vname(fb.vid != null ? fb.vid : mainVid) || 'τον κύριο προμηθευτή') + '.');
          }
        }
      }
      else pick = offers[mainVid] || (mainVid ? { vid: x.vid, cost: mainCost, disc: 0, net: mainCost, code: '' } : list[0]);
      if (!pick || !(pick.net > 0)) { skippedNoCost.push(x.nm); continue; }
      const main = offers[mainVid];
      if (suspicious) notes.push('Αγνοήθηκε κωδικός προμηθευτή με προφανώς λάθος τιμή/έκπτωση στο SimplyPOS.');
      if (o.cheapest && main && pick.net < 0.5 * main.net) notes.push('Πολύ χαμηλότερη τιμή από τον κύριο προμηθευτή — ελέγξτε μονάδα/τιμή.');
      if (cstk < 0) notes.push('Αρνητικό τρέχον απόθεμα.');
      if (!(pick.vid in meta.vendors) && !(String(pick.vid) in meta.vendors)) notes.push('Ο προμηθευτής δεν υπάρχει πλέον στη λίστα προμηθευτών.');
      const vat = (Number(x.vat) || 0) / 100;
      const net = r4(pick.net), val = r2(q * net), gross = r2(val * (1 + vat));
      const mnet = main ? r4(main.net) : (mainCost > 0 ? r4(mainCost) : null);
      rows.push({
        cd, scl: (x.sCdList || []).map((s) => [String(s.vid), s.code || '']),
        bcs: [...new Set([...(x.bCdList || []).map((b) => b.bCd), x.bCd || ''].filter((b) => /^\d{8,14}$/.test(String(b))))], pid: x.pid ? String(x.pid).toLowerCase() : '', pk: pk || '', boxes, qPre, nm: String(x.nm || '').trim(), br, cur, yAvg, yN, base, cstk, dstk, q, vat,
        vid: String(pick.vid), vn: vname(pick.vid), code: pick.code || (x.sCdList || []).find((s) => String(s.vid) === String(pick.vid))?.code || '',
        cost: pick.cost, disc: pick.disc / 100, net, val, gross,
        mvn: vname(mainVid), mnet, save: mnet == null ? null : r2((mnet - net) * q), note: notes.join(' ')
      });
    }
    const filtered = o.vendor ? rows.filter((r) => r.vid === String(o.vendor)) : rows;
    const tot = {}; filtered.forEach((r) => { tot[r.vn] = (tot[r.vn] || 0) + r.val; });
    const vorder = Object.keys(tot).sort((a, b) => tot[b] - tot[a]);
    filtered.sort((a, b) => vorder.indexOf(a.vn) - vorder.indexOf(b.vn) || b.q - a.q || a.nm.localeCompare(b.nm, 'el'));
    return { rows: filtered, vorder, tot, skippedNoCost, unknown };
  }

  // ---------- Excel ----------
  // Κανόνας: ποσότητες, τιμές, αξίες, σύνολα = ΤΕΛΙΚΟΙ ΑΡΙΘΜΟΙ (χωρίς υπολογισμούς).
  // Τύποι υπάρχουν ΜΟΝΟ για το κίτρινο κελί διαγραφής (απόκρυψη γραμμών) και τα σύνολα όσων φαίνονται.
  async function buildExcel(o, meta, res, periods) {
    if (!window.ExcelJS) await loadScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Εργαλείο Παραγγελιών SimplyPOS';
    wb.calcProperties.fullCalcOnLoad = true;
    const F = 'Arial', HDR = 'FF1F4E78', GREY = 'FFD9D9D9', TOT = 'FFDDEBF7', YEL = 'FFFFFF00';
    const EUR = '#,##0.00 "€"', EUR4 = '#,##0.0000 "€"', QTY = '#,##0', PCT = '0%';
    const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } }, border = { top: thin, left: thin, bottom: thin, right: thin };
    const med = { style: 'medium', color: { argb: 'FF000000' } };
    const stamp = new Date();
    const stampTxt = `Δημιουργήθηκε ${pad(stamp.getDate())}/${pad(stamp.getMonth() + 1)}/${stamp.getFullYear()} ${pad(stamp.getHours())}:${pad(stamp.getMinutes())} από το SimplyPOS (${meta.active.name}). Δεν καταχωρήθηκε καμία παραγγελία στο SimplyPOS.`;
    const header = (ws, rowNo, titles, widths) => {
      const row = ws.getRow(rowNo);
      titles.forEach((t, i) => {
        const c = row.getCell(i + 1); c.value = t;
        c.font = { name: F, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDR } };
        c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; c.border = border;
      });
      row.height = 45;
      widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    };
    const fillGrey = (c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY } }; };
    // Εκτύπωση: A4 όρθιο, προσαρμογή στο πλάτος, επανάληψη επικεφαλίδας, αρίθμηση σελίδων
    const printA4 = (ws, area, titleRow) => {
      ws.pageSetup = {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        horizontalCentered: true, printArea: area, printTitlesRow: titleRow ? `${titleRow}:${titleRow}` : undefined,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 }
      };
      ws.headerFooter = { oddFooter: '&L&8' + ws.name + '&R&8Σελίδα &P από &N' };
    };
    const baseTxt = o.avg2y ? 'Μ.Ο. πωλήσεων ίδιας περιόδου στα προηγούμενα έτη' : 'πωλήσεις περιόδου';
    const qtyRule = o.subStock ? `Ποσότητα = ${baseTxt} − τρέχον απόθεμα (στρογγυλοποίηση προς τα πάνω)` : `Ποσότητα = ${baseTxt} (χωρίς αφαίρεση αποθέματος, στρογγυλοποίηση προς τα πάνω)`;
    const printTxt = { summary: 'ΣΥΝΟΨΗ (φύλλο Σύνοψη)', all: 'ΣΥΝΟΛΙΚΗ (φύλλο Δεδομένα)', vendors: 'ΠΡΟΜΗΘΕΥΤΕΣ (ένα φύλλο ανά προμηθευτή)' };
    const settingsTxt = [
      ['Κατάστημα', meta.active.name],
      ['Περίοδος πωλήσεων', `${isoToGr(o.from)} – ${isoToGr(o.to)}`],
      ['Αφαίρεση τρέχοντος αποθέματος', o.subStock ? 'ΝΑΙ' : 'ΟΧΙ (ποσότητα = πωλήσεις)'],
      ['Μ.Ο. πωλήσεων ίδιας περιόδου (έως 5 προηγούμενα έτη)', o.avg2y ? `ΝΑΙ — ${periods.yrs.join(' · ')} · διαιρείται με όσα έτη είχε πωλήσεις το είδος` : 'ΟΧΙ'],
      ['Φθηνότερος προμηθευτής', o.cheapest ? 'ΝΑΙ' : (o.vendor ? 'ΟΧΙ — είδη με δηλωμένο τον προμηθευτή: ' + (meta.vendors[o.vendor] || o.vendor) : 'ΟΧΙ (κύριος προμηθευτής κάθε είδους)')],
      ['Παραγγελία μόνο αν πωλήσεις >', `${o.minSales} τεμ.`],
      ['Υπέρβαση επιθυμητού αποθέματος', o.overDesired ? 'ΝΑΙ' : 'ΟΧΙ (τρέχον + παραγγελία ≤ επιθυμητό)'],
      ['Προμηθευτής', o.vendor ? (meta.vendors[o.vendor] || o.vendor) : 'Όλοι'],
      ['Κατασκευαστής', o.brand ? (meta.brandName[o.brand] || o.brand) : 'Όλοι'],
      ['Στρογγυλοποίηση κιβωτίων', o.boxRound ? 'ΝΑΙ — στο πλησιέστερο κιβώτιο (πεδίο «Ποσότητα» στις Πληροφορίες του είδους· χωρίς τιμή = ως έχει)' : 'ΟΧΙ'],
      ['Εκτύπωση', printTxt[o.print] || printTxt.vendors]
    ];
    const SM_DEL = 13, SM_HDR = 17, SM_S = 18;           // γραμμές Σύνοψης
    const THR = `Σύνοψη!$B$${SM_DEL}`;                     // κίτρινο κελί διαγραφής
    const sm = wb.addWorksheet('Σύνοψη', { views: [{ state: 'frozen', ySplit: SM_HDR }] });
    const dt = wb.addWorksheet('Δεδομένα', { views: [{ state: 'frozen', xSplit: 4, ySplit: 4 }] });

    // ---- Δεδομένα (όλα τελικοί αριθμοί) — οι 4 πρώτες στήλες είναι αυτές που τυπώνονται ----
    // A Κωδ.προμ · B Είδος · C Ποσότητα · D Σύνολο με ΦΠΑ · E Προμηθευτής · F Κωδ.SimplyPOS · G-I Πωλήσεις · J Βάση
    // K Τρέχον · L Επιθυμητό · M Τιμή · N Έκπτωση · O Καθ.κόστος · P Καθ.αξία · Q ΦΠΑ · R Κύριος · S Κόστος κύριου
    // T Εξοικονόμηση · U Σημείωση · V Κατασκευαστής · W Α/Α (φίλτρο) · X Κλειδί (φίλτρο)
    dt.getCell('A1').value = `Παραγγελίες — ${meta.active.name} — πωλήσεις ${isoToGr(o.from)}–${isoToGr(o.to)}`;
    dt.getCell('A1').font = { name: F, bold: true, size: 14 };
    dt.getCell('A2').value = 'Όλοι οι αριθμοί είναι τελικοί (χωρίς υπολογισμούς). Γκρι στήλη C = ποσότητα παραγγελίας. Διαγραφή γραμμών: κίτρινο κελί στη «Σύνοψη».';
    dt.getCell('A2').font = { name: F, italic: true };
    dt.getCell('A3').value = stampTxt;
    dt.getCell('A3').font = { name: F, size: 9, color: { argb: 'FF595959' } };
    header(dt, 4, ['Κωδ. προμηθευτή', 'Είδος', 'Ποσότητα παραγγελίας', 'Σύνολο με ΦΠΑ', 'Προμηθευτής', 'Κωδ. SimplyPOS',
      `Πωλήσεις ${periods.cur}`, 'Μ.Ο. ίδιας περιόδου προηγ. ετών', 'Έτη με πωλήσεις (από 5)', 'Βάση (τεμ.)',
      'Τρέχον απόθεμα', 'Επιθυμητό απόθεμα', 'Τιμή καταλόγου', 'Έκπτωση', 'Καθαρό κόστος', 'Καθαρή αξία', 'ΦΠΑ',
      'Κύριος προμηθευτής', 'Καθαρό κόστος κύριου', 'Εξοικονόμηση έναντι κύριου', 'Σημείωση', 'Κατασκευαστής', 'Α/Α (φίλτρο)', 'Κλειδί (φίλτρο)',
      'Τεμ. ανά κιβώτιο', 'Κιβώτια', 'Ποσότητα πριν τη στρογγυλοποίηση'],
      [15, 46, 11, 13, 24, 10, 11, 11, 11, 10, 10, 10, 12, 9, 12, 12, 7, 24, 12, 13, 50, 26, 8, 10, 10, 9, 13]);
    const S = 5, E = Math.max(S, S + res.rows.length - 1);
    const BLACK = { name: F }, HELP = { name: F, size: 8, color: { argb: 'FF808080' } };
    res.rows.forEach((r, i) => {
      const n = S + i;
      const vals = [r.code, r.nm, r.q, r.gross, r.vn, r.cd, r.cur, o.avg2y ? r4(r.yAvg) : '', o.avg2y ? r.yN : '', r4(r.base), r.cstk, r.dstk, r4(r.cost), r.disc, r.net,
        r.val, r.vat, r.mvn, r.mnet == null ? '' : r.mnet, r.save == null ? '' : r.save, r.note, r.br || ''];
      const row = dt.getRow(n);
      vals.forEach((v, j) => { const c = row.getCell(j + 1); c.value = v; c.border = border; c.font = BLACK; });
      const qc = row.getCell(3); fillGrey(qc); qc.font = { name: F, bold: true };
      [[3, QTY], [4, EUR], [7, QTY], [8, '#,##0.0'], [9, QTY], [10, '#,##0.0'], [11, QTY], [12, QTY], [13, EUR4], [14, PCT], [15, EUR4], [16, EUR], [17, PCT], [19, EUR4], [20, EUR]]
        .forEach(([j, f]) => { row.getCell(j).numFmt = f; });
      const nc = row.getCell(21); nc.alignment = { wrapText: true, vertical: 'top' };
      if (r.note) nc.font = { name: F, size: 9, color: { argb: 'FFC00000' } };
      const rank = res.rows.slice(0, i + 1).filter((x) => x.vn === r.vn).length;
      row.getCell(23).value = { formula: `IF(C${n}>${THR},COUNTIFS($E$${S}:E${n},E${n},$C$${S}:C${n},">"&${THR}),"")`, result: rank };
      row.getCell(24).value = { formula: `IF(W${n}="","",E${n}&"|"&W${n})`, result: `${r.vn}|${rank}` };
      row.getCell(23).font = HELP; row.getCell(24).font = HELP;
      [[25, r.pk], [26, r.boxes == null ? '' : r.boxes], [27, r.qPre]].forEach(([j, v]) => { const c = row.getCell(j); c.value = v; c.border = border; c.font = BLACK; c.numFmt = QTY; });
    });
    const tRow = dt.getRow(E + 1);
    const sumRes = (k) => res.rows.reduce((a, r) => a + k(r), 0);
    if (res.rows.length) {
      tRow.getCell(1).value = 'ΣΥΝΟΛΟ';
      tRow.getCell(2).value = 'Γραμμές: ' + res.rows.length;
      tRow.getCell(3).value = sumRes((r) => r.q);
      tRow.getCell(4).value = r2(sumRes((r) => r.gross));
      tRow.getCell(16).value = r2(sumRes((r) => r.val));
      tRow.getCell(20).value = r2(sumRes((r) => r.save || 0));
      for (let j = 1; j <= 22; j++) { const c = tRow.getCell(j); c.font = { name: F, bold: true }; c.border = border; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOT } }; }
      fillGrey(tRow.getCell(3));
      [[3, QTY], [4, EUR], [16, EUR], [20, EUR]].forEach(([j, f]) => { tRow.getCell(j).numFmt = f; });
      dt.autoFilter = { from: { row: 4, column: 1 }, to: { row: E, column: 22 } };
    }
    printA4(dt, `A1:D${E + 1}`, 4);
    const D = (c) => `Δεδομένα!$${c}$${S}:$${c}$${E}`;

    // ---- φύλλα προμηθευτών (γραμμές φιλτράρονται από το κίτρινο κελί) ----
    // A Κωδ.προμ · B Είδος · C Ποσότητα · D Σύνολο με ΦΠΑ (τυπώνονται) · E Καθ.κόστος · F Καθ.αξία · G ΦΠΑ · H Κατασκευαστής
    const used = new Set(['Σύνοψη', 'Δεδομένα']); const sheetOf = {}; const vendorSheets = [];
    res.vorder.forEach((vn) => {
      let nm = vn.replace(/[\\\/\?\*\[\]:']/g, '').slice(0, 28).trim() || 'Προμηθευτής', base = nm, k = 2;
      while (used.has(nm)) nm = base.slice(0, 25) + ' ' + (k++);
      used.add(nm); sheetOf[vn] = nm;
      const ws = wb.addWorksheet(nm, { views: [{ state: 'frozen', ySplit: 5 }] });
      vendorSheets.push(ws);
      ws.getCell('A1').value = `Παραγγελία προς ${vn} — ${meta.active.name}`; ws.getCell('A1').font = { name: F, bold: true, size: 14 };
      ws.getCell('A2').value = `Πωλήσεις ${isoToGr(o.from)}–${isoToGr(o.to)}. ${qtyRule}.`;
      ws.getCell('A2').font = { name: F, italic: true };
      ws.getCell('A3').value = stampTxt; ws.getCell('A3').font = { name: F, size: 9, color: { argb: 'FF595959' } };
      ws.getCell('A4').value = 'Προμηθευτής:'; ws.getCell('A4').font = { name: F, bold: true };
      ws.getCell('B4').value = vn; ws.getCell('B4').font = { name: F, bold: true };
      header(ws, 5, ['Κωδ. προμηθευτή', 'Είδος', 'Ποσότητα παραγγελίας', 'Σύνολο με ΦΠΑ', 'Καθαρό κόστος', 'Καθαρή αξία', 'ΦΠΑ', 'Κατασκευαστής', 'Κιβώτια'], [15, 50, 12, 14, 13, 13, 7, 26, 9]);
      const rs = res.rows.filter((r) => r.vn === vn);
      const cnt = `COUNTIFS(${D('E')},$B$4,${D('C')},">"&${THR})`;
      rs.forEach((r, k2) => {
        const n = 6 + k2, kk = k2 + 1, row = ws.getRow(n);
        const M = `MATCH($B$4&"|"&${kk},${D('X')},0)`;
        const g = (col, extra) => `IF(${kk}<=${cnt},INDEX(${D(col)},${M})${extra || ''},"")`;
        const vals = [
          { formula: g('A', '&""'), result: String(r.code || '') }, { formula: g('B'), result: r.nm },
          { formula: g('C'), result: r.q }, { formula: g('D'), result: r.gross },
          { formula: g('O'), result: r.net }, { formula: g('P'), result: r.val },
          { formula: g('Q'), result: r.vat }, { formula: g('V', '&""'), result: r.br || '' },
          { formula: g('Z', '&""'), result: r.boxes == null ? '' : String(r.boxes) }
        ];
        vals.forEach((v, j) => { const c = row.getCell(j + 1); c.value = v; c.border = border; c.font = BLACK; });
        fillGrey(row.getCell(3)); row.getCell(3).font = { name: F, bold: true };
        [[3, QTY], [4, EUR], [5, EUR4], [6, EUR], [7, PCT]].forEach(([j, f]) => { row.getCell(j).numFmt = f; });
      });
      const e = 6 + Math.max(rs.length, 1) - 1, tr = ws.getRow(e + 1);
      tr.getCell(1).value = 'ΣΥΝΟΛΟ';
      tr.getCell(2).value = { formula: `"Γραμμές: "&${cnt}`, result: 'Γραμμές: ' + rs.length };
      tr.getCell(3).value = { formula: `SUM(C6:C${e})`, result: rs.reduce((a, r) => a + r.q, 0) };
      tr.getCell(4).value = { formula: `SUM(D6:D${e})`, result: r2(rs.reduce((a, r) => a + r.gross, 0)) };
      tr.getCell(6).value = { formula: `SUM(F6:F${e})`, result: r2(rs.reduce((a, r) => a + r.val, 0)) };
      for (let j = 1; j <= 8; j++) { const c = tr.getCell(j); c.font = { name: F, bold: true }; c.border = border; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOT } }; }
      fillGrey(tr.getCell(3));
      [[3, QTY], [4, EUR], [6, EUR]].forEach(([j, f]) => { tr.getCell(j).numFmt = f; });
      printA4(ws, `A1:D${e + 1}`, 5);
    });

    // ---- Σύνοψη ----
    sm.getCell('A1').value = `Παραγγελίες — ${meta.active.name}`; sm.getCell('A1').font = { name: F, bold: true, size: 14 };
    settingsTxt.forEach(([k, v], i) => {
      const r = sm.getRow(2 + i); r.getCell(1).value = k; r.getCell(1).font = { name: F, bold: true };
      r.getCell(2).value = v; r.getCell(2).font = { name: F };
    });
    const dl = sm.getCell(`A${SM_DEL}`); dl.value = 'ΔΙΑΓΡΑΦΗ γραμμών με ποσότητα παραγγελίας έως:'; dl.font = { name: F, bold: true };
    const dc = sm.getCell(`B${SM_DEL}`);
    dc.value = 0; dc.font = { name: F, bold: true, size: 12, color: { argb: 'FF0000FF' } };
    dc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: YEL } };
    dc.alignment = { horizontal: 'center' }; dc.border = { top: med, left: med, bottom: med, right: med };
    dc.dataValidation = { type: 'list', allowBlank: false, formulae: ['"0,1,2,3,4,5,10,20,50"'], showErrorMessage: false,
      showInputMessage: true, promptTitle: 'Διαγραφή', prompt: 'Αφαιρούνται οι γραμμές με ποσότητα παραγγελίας μικρότερη ή ίση με αυτή την τιμή (σε όλα τα φύλλα προμηθευτών). Μπορείτε να γράψετε και άλλο αριθμό.' };
    sm.getCell(`C${SM_DEL}`).value = '◄ επιλέξτε (0 = όλα)'; sm.getCell(`C${SM_DEL}`).font = { name: F, size: 9, color: { argb: 'FF595959' } };
    sm.getCell(`A${SM_DEL + 1}`).value = `${qtyRule}. Καθαρό κόστος = τιμή × (1 − έκπτωση). Όλα τελικοί αριθμοί, χωρίς υπολογισμούς.`;
    sm.getCell(`A${SM_DEL + 1}`).font = { name: F, italic: true, size: 10 };
    sm.getCell(`A${SM_DEL + 2}`).value = stampTxt; sm.getCell(`A${SM_DEL + 2}`).font = { name: F, size: 9, color: { argb: 'FF595959' } };
    header(sm, SM_HDR, ['Προμηθευτής', 'Γραμμές', 'Ποσότητα παραγγελίας', 'Σύνολο με ΦΠΑ', 'Καθαρή αξία', 'Εξοικονόμηση έναντι κύριου', 'Φύλλο'], [34, 30, 13, 15, 14, 16, 30]);
    const T = `$B$${SM_DEL}`;
    res.vorder.forEach((vn, i) => {
      const n = SM_S + i, row = sm.getRow(n), rs = res.rows.filter((r) => r.vn === vn);
      const cond = `${D('E')},A${n},${D('C')},">"&${T}`;
      row.getCell(1).value = vn; row.getCell(1).font = BLACK;
      row.getCell(2).value = { formula: `COUNTIFS(${cond})`, result: rs.length };
      row.getCell(3).value = { formula: `SUMIFS(${D('C')},${cond})`, result: rs.reduce((a, r) => a + r.q, 0) };
      row.getCell(4).value = { formula: `SUMIFS(${D('D')},${cond})`, result: r2(rs.reduce((a, r) => a + r.gross, 0)) };
      row.getCell(5).value = { formula: `SUMIFS(${D('P')},${cond})`, result: r2(rs.reduce((a, r) => a + r.val, 0)) };
      row.getCell(6).value = { formula: `SUMIFS(${D('T')},${cond})`, result: r2(rs.reduce((a, r) => a + (r.save || 0), 0)) };
      row.getCell(7).value = { text: sheetOf[vn], hyperlink: `#'${sheetOf[vn]}'!A1` };
      row.getCell(7).font = { name: F, color: { argb: 'FF0563C1' }, underline: true };
      for (let j = 1; j <= 7; j++) row.getCell(j).border = border;
      for (let j = 2; j <= 6; j++) row.getCell(j).font = BLACK;
      fillGrey(row.getCell(3)); row.getCell(3).font = { name: F, bold: true };
      [[2, QTY], [3, QTY], [4, EUR], [5, EUR], [6, EUR]].forEach(([j, f]) => { row.getCell(j).numFmt = f; });
    });
    const e3 = SM_S + Math.max(res.vorder.length, 1) - 1, tr3 = sm.getRow(e3 + 1);
    tr3.getCell(1).value = 'ΣΥΝΟΛΟ';
    [['B', 2, res.rows.length], ['C', 3, sumRes((r) => r.q)], ['D', 4, r2(sumRes((r) => r.gross))],
     ['E', 5, r2(sumRes((r) => r.val))], ['F', 6, r2(sumRes((r) => r.save || 0))]]
      .forEach(([c, j, v]) => { tr3.getCell(j).value = { formula: `SUM(${c}${SM_S}:${c}${e3})`, result: v }; });
    for (let j = 1; j <= 7; j++) { const c = tr3.getCell(j); c.font = { name: F, bold: true }; c.border = border; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOT } }; }
    fillGrey(tr3.getCell(3));
    [[2, QTY], [3, QTY], [4, EUR], [5, EUR], [6, EUR]].forEach(([j, f]) => { tr3.getCell(j).numFmt = f; });
    let nr = e3 + 3;
    const note = (t, b) => { const c = sm.getCell(`A${nr++}`); c.value = t; c.font = { name: F, size: b ? 11 : 10, bold: !!b }; };
    note('Σημειώσεις', true);
    note('• Είδη με ποσότητα παραγγελίας 0 δεν περιλαμβάνονται.');
    note('• Οι γραμμές με κόκκινη «Σημείωση» στο φύλλο Δεδομένα θέλουν έλεγχο.');
    if (res.skippedNoCost.length) note(`• Παραλείφθηκαν ${res.skippedNoCost.length} είδη χωρίς κόστος αγοράς: ${res.skippedNoCost.slice(0, 15).join(', ')}${res.skippedNoCost.length > 15 ? '…' : ''}`);
    if (res.unknown.length) note(`• ${res.unknown.length} κωδικοί με πωλήσεις δεν βρέθηκαν στη λίστα ειδών αγορών.`);
    if (o.boxRound) note(`• Στρογγυλοποίηση κιβωτίων: ${res.rows.filter((r) => r.pk).length} από ${res.rows.length} είδη έχουν «Ποσότητα» κιβωτίου στο SimplyPOS και στρογγυλοποιήθηκαν· τα υπόλοιπα έμειναν ως έχουν.`);
    note('• Το καθαρό κόστος δεν περιλαμβάνει ελάχιστη παραγγελία ή μεταφορικά.');
    note(`• Εκτύπωση: ${printTxt[o.print] || printTxt.vendors} — A4, μόνο οι στήλες Κωδ. προμηθευτή, Είδος, Ποσότητα παραγγελίας, Σύνολο με ΦΠΑ.`);
    printA4(sm, `A1:F${nr - 1}`, SM_HDR);

    // ---- ποιο φύλλο/φύλλα είναι επιλεγμένα στο άνοιγμα → αυτά τυπώνει το Ctrl+P ----
    const mode = o.print || 'vendors';
    const sel = mode === 'summary' ? [sm] : mode === 'all' ? [dt] : (vendorSheets.length ? vendorSheets : [sm]);
    wb.worksheets.forEach((ws) => { ws.views = [Object.assign({}, ws.views && ws.views[0], { tabSelected: sel.includes(ws) })]; });
    const activeIdx = wb.worksheets.indexOf(sel[0]);
    wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: activeIdx, visibility: 'visible' }];
    return wb;
  }

  // ---------- UI ----------
  const css = `
  #${PID}{position:fixed;top:16px;right:16px;z-index:2147483646;width:380px;max-height:calc(100vh - 32px);display:flex;flex-direction:column;
    background:#fff;color:#1f2933;border-radius:14px;box-shadow:0 12px 40px rgba(15,35,60,.28);font:14px/1.4 Arial,Helvetica,sans-serif;overflow:hidden}
  #${PID} *{box-sizing:border-box;font-family:inherit}
  #${PID} .hd{background:#1F4E78;color:#fff;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;cursor:move}
  #${PID} .hd b{font-size:16px;letter-spacing:.2px}
  #${PID} .hd button{background:transparent;border:0;color:#fff;font-size:20px;cursor:pointer;line-height:1;padding:2px 6px}
  #${PID} .bd{padding:14px 16px;overflow:auto}
  #${PID} .st{background:#eef4fa;border-radius:8px;padding:8px 10px;margin-bottom:12px;font-size:13px}
  #${PID} .fld{margin-bottom:12px}
  #${PID} .lbl{display:block;font-weight:bold;margin-bottom:5px;font-size:13px}
  #${PID} .num{display:inline-block;width:20px;height:20px;border-radius:50%;background:#1F4E78;color:#fff;text-align:center;font-size:12px;line-height:20px;margin-right:6px}
  #${PID} .row2{display:flex;gap:8px}
  #${PID} input[type=date],#${PID} input[type=number],#${PID} input[type=text],#${PID} select{width:100%;padding:8px;border:1px solid #c5ced8;border-radius:8px;font-size:14px;background:#fff;color:#1f2933}
  #${PID} .seg{display:flex;border:1px solid #c5ced8;border-radius:8px;overflow:hidden}
  #${PID} .seg button{flex:1;padding:8px;border:0;background:#fff;cursor:pointer;font-weight:bold;color:#52606d}
  #${PID} .seg button.on{background:#1F4E78;color:#fff}
  #${PID} .hint{font-size:11.5px;color:#616e7c;margin-top:4px}
  #${PID} .go{width:100%;padding:12px;border:0;border-radius:10px;background:#2e7d32;color:#fff;font-size:15px;font-weight:bold;cursor:pointer;margin-top:4px}
  #${PID} .go:disabled{background:#9aa5b1;cursor:wait}
  #${PID} .msg{margin-top:12px;font-size:13px}
  #${PID} .msg.err{color:#b71c1c}
  #${PID} table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12.5px}
  #${PID} td{padding:4px 2px;border-bottom:1px solid #e4e7eb}
  #${PID} td.r{text-align:right;white-space:nowrap}
  #${PID} .big{font-size:15px;font-weight:bold}
  #${PID} .quick{background:#f0f7f1;border:1px solid #cfe5d3;border-radius:10px;padding:10px;margin-bottom:12px}
  #${PID} .qb{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}
  #${PID} .qb button{flex:1 1 45%;padding:12px 6px;border:0;border-radius:10px;background:#2e7d32;color:#fff;font-weight:bold;font-size:14px;cursor:pointer}
  #${PID} .qb button[data-q=excel]{background:#1F4E78}
  #${PID} .qb button:disabled{background:#9aa5b1;cursor:wait}
  #${PID} details summary{cursor:pointer;font-weight:bold;padding:8px 0;color:#1F4E78}
  #${PID} .ft{font-size:11px;color:#7b8794;padding:8px 16px;border-top:1px solid #e4e7eb}`;
  const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);

  // προεπιλογή: τελευταίες 7 ημέρες μέχρι χθες
  const today = new Date(); const y = new Date(today); y.setDate(y.getDate() - 1); const f0 = new Date(y); f0.setDate(f0.getDate() - 6);
  const saved = store.get('opts', {});
  const panel = document.createElement('div'); panel.id = PID;
  const seg = (key, val) => `<div class="seg" data-key="${key}"><button type="button" data-v="1" class="${val ? 'on' : ''}">ΝΑΙ</button><button type="button" data-v="0" class="${val ? '' : 'on'}">ΟΧΙ</button></div>`;
  panel.innerHTML = `
    <div class="hd"><b>📦 Παραγγελίες</b><button type="button" title="Κλείσιμο" data-act="close">×</button></div>
    <div class="bd">
      <div class="st" id="spo-store">Φόρτωση στοιχείων SimplyPOS…</div>
      <div class="quick"><div class="lbl">⚡ Γρήγορη παραγγελία — πωλήσεις τελευταίων 7 ημερών</div>
        <div class="qb"><button type="button" data-q="ragousis" disabled>🛒 ΡΑΓΚΟΥΣΗΣ</button><button type="button" data-q="mmm" disabled>🛒 Mmmgroup</button><button type="button" data-q="mas" disabled>🛒 ΜΑΣΟΥΤΗΣ</button>
        <button type="button" data-q="zoumis" disabled>✉️ ΖΟΥΜΗΣ</button><button type="button" data-q="michalios" disabled>✉️ ΜΙΧΑΛΙΟΣ</button><button type="button" data-q="mietsel" disabled>✉️ ΜΙΕΤΣΕΛ</button><button type="button" data-q="marinopoulos" disabled>✉️ ΜΑΡΙΝΟΠΟΥΛΟΣ</button><button type="button" data-q="stefanidis" disabled>✉️ ΣΤΕΦΑΝΙΔΗΣ</button><button type="button" data-q="asteras" disabled>✉️ ASTERAS</button><button type="button" data-q="excel" disabled>📄 Μόνο Excel</button></div>
        <div class="hint">Ένα πάτημα: φτιάχνει το Excel και ανοίγει το site του προμηθευτή. Εκεί πατάτε <b>μία φορά</b> τον σελιδοδείκτη και η παραγγελία μπαίνει στο καλάθι. <b>Κουμπιά ✉️</b>: η παραγγελία πάει με e-mail από rrms2021@gmail.com (με αυτόματη αποστολή: αμέσως· αλλιώς ανοίγει έτοιμο μήνυμα στο Gmail). «Μόνο Excel» = όλοι οι προμηθευτές.</div></div>
      <details id="spo-adv"><summary>⚙️ Ρυθμίσεις (για προχωρημένους)</summary>
      <div class="fld" style="margin-top:10px"><span class="lbl"><span class="num">1</span>Ημερομηνίες πωλήσεων (από – έως)</span>
        <div class="row2"><input type="date" id="spo-from" value="${isoOf(f0)}"><input type="date" id="spo-to" value="${isoOf(y)}"></div>
        <div class="row2" style="margin-top:6px;gap:6px">
          ${[7, 10, 14, 30].map((d) => `<button type="button" class="spo-q" data-d="${d}" style="flex:1;padding:5px;border:1px solid #c5ced8;border-radius:6px;background:#f5f7fa;cursor:pointer;font-size:12px">${d} ημ.</button>`).join('')}
        </div></div>
      <div class="fld"><span class="lbl"><span class="num">2</span>Αφαίρεση τρέχοντος αποθέματος</span>${seg('subStock', saved.subStock ?? false)}
        <div class="hint">ΟΧΙ: ποσότητα = πωλήσεις περιόδου. ΝΑΙ: ποσότητα = πωλήσεις − τρέχον απόθεμα.</div></div>
      <div class="fld"><span class="lbl"><span class="num">3</span>Φθηνότερος προμηθευτής</span>${seg('cheapest', saved.cheapest ?? true)}
        <div class="hint">ΝΑΙ: κάθε είδος στον προμηθευτή με το χαμηλότερο καθαρό κόστος. ΟΧΙ: τα είδη που έχουν δηλωμένο στην καρτέλα «Προμηθευτής» τον προμηθευτή που διαλέγετε παρακάτω· αν δεν διαλέξετε, στον κύριο προμηθευτή κάθε είδους.</div></div>
      <div class="fld"><span class="lbl"><span class="num">4</span>Παραγγελία μόνο αν οι πωλήσεις ξεπερνούν (τεμ.)</span>
        <input type="number" id="spo-min" min="0" step="1" value="${saved.minSales ?? 0}">
        <div class="hint">0 = όλα τα είδη που πουλήθηκαν.</div></div>
      <div class="fld"><span class="lbl"><span class="num">5</span>Μ.Ο. πωλήσεων ίδιας περιόδου — έως 5 προηγούμενα έτη</span>${seg('avg2y', saved.avg2y ?? false)}
        <div class="hint" id="spo-avg-hint">ΝΑΙ: βάση = μέσος όρος πωλήσεων των ίδιων ημερομηνιών στα 5 προηγούμενα έτη. Διαιρείται μόνο με τα έτη που το είδος είχε πωλήσεις (π.χ. πωλήσεις σε 3 έτη → Μ.Ο. 3 ετών).</div></div>
      <div class="fld"><span class="lbl"><span class="num">6</span>Υπέρβαση επιθυμητού αποθέματος</span>${seg('overDesired', saved.overDesired ?? true)}
        <div class="hint">ΟΧΙ: η παραγγελία κόβεται ώστε τρέχον + παραγγελία ≤ επιθυμητό απόθεμα.</div></div>
      <div class="fld"><span class="lbl">Προμηθευτής (προαιρετικό)</span><select id="spo-vendor"><option value="">Όλοι οι προμηθευτές</option></select>
        <div class="hint">Με «Φθηνότερος = ΟΧΙ»: όλα τα είδη που έχουν αυτόν τον προμηθευτή δηλωμένο στο SimplyPOS (κύριο ή δευτερεύοντα). Με «ΝΑΙ»: μόνο όσα είναι φθηνότερα σε αυτόν.</div></div>
      <div class="fld"><span class="lbl"><span class="num">7</span>Εκτύπωση (A4)</span><select id="spo-print">
          <option value="summary">1. ΣΥΝΟΨΗ — φύλλο Σύνοψη</option>
          <option value="all">2. ΣΥΝΟΛΙΚΗ — φύλλο Δεδομένα (όλα τα είδη)</option>
          <option value="vendors">3. ΠΡΟΜΗΘΕΥΤΕΣ — μία σελίδα για κάθε προμηθευτή</option></select>
        <div class="hint">Το Excel ανοίγει με αυτό επιλεγμένο: Ctrl+P → τυπώνει σε A4 μόνο Κωδ. προμηθευτή, Είδος, Ποσότητα παραγγελίας, Σύνολο με ΦΠΑ.</div></div>
      <div class="fld"><span class="lbl"><span class="num">8</span>Στρογγυλοποίηση κιβωτίων</span>${seg('boxRound', saved.boxRound ?? false)}
        <div class="hint">ΝΑΙ: η ποσότητα πάει στο πλησιέστερο κιβώτιο (τουλάχιστον 1) με βάση το πεδίο «Ποσότητα» στις Πληροφορίες του είδους. Είδη χωρίς «Ποσότητα» μένουν ως έχουν.</div></div>
      <div class="fld"><span class="lbl"><span class="num">9</span>Καταχώρηση παραγγελίας</span><select id="spo-send">
          <option value="">ΟΧΙ — μόνο Excel</option><option value="ragousis">ΡΑΓΚΟΥΣΗΣ — καλάθι στο site</option><option value="mmm">Mmmgroup — καλάθι στο site</option><option value="mas">ΜΑΣΟΥΤΗΣ — καλάθι στο site</option><option value="zoumis">ΖΟΥΜΗΣ — e-mail (cavazoumis@gmail.com)</option><option value="michalios">ΜΙΧΑΛΙΟΣ — e-mail (mixalioscustomers@gmail.com)</option><option value="mietsel">ΜΙΕΤΣΕΛ — e-mail (paros.mietsel@gmail.com)</option><option value="marinopoulos">ΜΑΡΙΝΟΠΟΥΛΟΣ ΙΚΕ — e-mail (marinopl@yahoo.com)</option><option value="stefanidis">ΣΤΕΦΑΝΙΔΗΣ Ι Α.Ε. — e-mail (info@stefanidisgroup.gr)</option><option value="asteras">ASTERAS — e-mail (ktimaroussoswinery@gmail.com)</option></select>
        <div class="hint">Site: μετά το Excel πατήστε «Άνοιγμα …» και στο site του προμηθευτή πατήστε ξανά αυτόν τον σελιδοδείκτη· μπαίνουν στο καλάθι όλα τα είδη του Excel. E-mail: ανοίγει έτοιμο μήνυμα στο Gmail. Την αποστολή την κάνετε πάντα εσείς.</div></div>
      <div class="fld"><span class="lbl">✉️ Αυτόματη αποστολή e-mail (κουμπιά ✉️)</span>
        <div class="row2"><input type="text" id="spo-hook" placeholder="https://script.google.com/macros/s/…/exec"><button type="button" id="spo-hook-save" style="padding:8px 12px;border:0;border-radius:8px;background:#1F4E78;color:#fff;font-weight:bold;cursor:pointer">OK</button></div>
        <div class="hint" id="spo-hook-st"></div></div>
      <div class="fld"><span class="lbl">Κατασκευαστής (προαιρετικό)</span><select id="spo-brand"><option value="">Όλοι οι κατασκευαστές</option></select>
        <div class="hint">Μόνο είδη αυτού του κατασκευαστή (όπως είναι δηλωμένος στο είδος στο SimplyPOS).</div></div>
      <button type="button" class="go" id="spo-go" disabled>ΔΗΜΙΟΥΡΓΙΑ EXCEL</button></details>
      <div class="msg" id="spo-msg"></div>
    </div>
    <div class="ft">Μόνο ανάγνωση — δεν καταχωρείται τίποτα στο SimplyPOS. Excel με τελικούς αριθμούς, γκρι στήλη ποσότητας και κίτρινο κελί διαγραφής στη «Σύνοψη». Για άλλο κατάστημα, αλλάξτε το κατάστημα στο SimplyPOS και ανοίξτε ξανά.</div>`;
  document.body.appendChild(panel);

  const opts = () => {
    const segVal = (k) => panel.querySelector(`.seg[data-key="${k}"] button.on`)?.dataset.v === '1';
    return { from: $id('spo-from').value, to: $id('spo-to').value, subStock: segVal('subStock'), cheapest: segVal('cheapest'), avg2y: segVal('avg2y'),
             overDesired: segVal('overDesired'), minSales: Math.max(0, Number($id('spo-min').value) || 0), vendor: $id('spo-vendor').value,
             brand: $id('spo-brand').value, print: $id('spo-print').value, boxRound: segVal('boxRound'), send: $id('spo-send').value };
  };
  panel.addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.act === 'close') { panel.style.display = 'none'; return; }
    if (b.parentElement.classList.contains('seg')) { [...b.parentElement.children].forEach((x) => x.classList.toggle('on', x === b)); }
    if (b.classList.contains('spo-q')) {
      const d = Number(b.dataset.d), t = new Date(); t.setDate(t.getDate() - 1); const f = new Date(t); f.setDate(f.getDate() - (d - 1));
      $id('spo-from').value = isoOf(f); $id('spo-to').value = isoOf(t);
    }
  });
  // drag
  (() => { const hd = panel.querySelector('.hd'); let sx, sy, ox, oy, on = false;
    hd.addEventListener('mousedown', (e) => { if (e.target.tagName === 'BUTTON') return; on = true; sx = e.clientX; sy = e.clientY; const r = panel.getBoundingClientRect(); ox = r.left; oy = r.top; e.preventDefault(); });
    window.addEventListener('mousemove', (e) => { if (!on) return; panel.style.left = (ox + e.clientX - sx) + 'px'; panel.style.top = Math.max(0, oy + e.clientY - sy) + 'px'; panel.style.right = 'auto'; });
    window.addEventListener('mouseup', () => { on = false; }); })();

  const msg = (html, err) => { const m = $id('spo-msg'); m.className = 'msg' + (err ? ' err' : ''); m.innerHTML = html; };
  let meta;
  try {
    meta = await loadMeta();
    if (!meta.active) throw new Error('Δεν βρέθηκε κατάστημα. Είστε συνδεδεμένοι στο SimplyPOS;');
    $id('spo-store').innerHTML = `Κατάστημα: <b>${esc(meta.active.name)}</b>`;
    const sel = $id('spo-vendor');
    Object.entries(meta.vendors).sort((a, b) => a[1].localeCompare(b[1], 'el')).forEach(([id, nm]) => {
      const o = document.createElement('option'); o.value = id; o.textContent = nm; sel.appendChild(o);
    });
    if (saved.vendor && meta.vendors[saved.vendor]) sel.value = saved.vendor;
    const bsel = $id('spo-brand');
    [...meta.brands].sort((a, b) => a.name.localeCompare(b.name, 'el', { numeric: true })).forEach((b) => {
      const o = document.createElement('option'); o.value = b.id; o.textContent = b.name; bsel.appendChild(o);
    });
    if (saved.brand && meta.brandName[saved.brand]) bsel.value = saved.brand;
    if (SEND[saved.send]) $id('spo-send').value = saved.send;
    if (['summary', 'all', 'vendors'].includes(saved.print)) $id('spo-print').value = saved.print; else $id('spo-print').value = 'vendors';
    $id('spo-go').disabled = false; panel.querySelectorAll('.qb button').forEach((b) => { b.disabled = false; });
  } catch (e) {
    $id('spo-store').textContent = 'Σφάλμα φόρτωσης';
    msg(esc(e.message || e) + '<br>Βεβαιωθείτε ότι είστε συνδεδεμένοι στο app.simplypos.com.', true);
    return;
  }

  async function run(o, download) {
    if (!o.from || !o.to || o.from > o.to) throw new Error('Ελέγξτε τις ημερομηνίες (από ≤ έως).');
    const NY = 5;   // Μ.Ο.: έως 5 προηγούμενα έτη
    const yrs = Array.from({ length: NY }, (_, i) => [isoShiftYear(o.from, -(i + 1)), isoShiftYear(o.to, -(i + 1))]);
    const periods = { cur: `${isoToGr(o.from)}–${isoToGr(o.to)}`, yrs: yrs.map((y) => `${isoToGr(y[0])}–${isoToGr(y[1])}`) };
    o.yearsLbl = yrs.map((y) => y[0].slice(0, 4) === y[1].slice(0, 4) ? y[0].slice(0, 4) : y[0].slice(0, 4) + '/' + y[1].slice(2, 4));
    const bid = o.brand || '';   // id κατασκευαστή
    const [P, sCur, bIds, ...sYs] = await Promise.all([
      purchaseProducts(meta.active.id), salesByCode(o.from, o.to, meta.active.id),
      bid ? brandProductIds(bid) : Promise.resolve(null),
      ...(o.avg2y ? yrs.map((y) => salesByCode(y[0], y[1], meta.active.id)) : [])
    ]);
    // κατασκευαστής ανά είδος: χωρίς φίλτρο = αυτός που δείχνει η αναφορά· με φίλτρο = ο επιλεγμένος
    const brandOf = Object.assign({}, ...sYs.slice().reverse().map((s) => s.b), sCur.b);
    if (bid) {
      const allowed = new Set(P.filter((x) => x.pid && bIds.has(String(x.pid).toLowerCase())).map((x) => String(x.cd)));
      [sCur, ...sYs].forEach((s) => Object.keys(s.q).forEach((k) => { if (!allowed.has(k)) delete s.q[k]; }));
      allowed.forEach((k) => { brandOf[k] = meta.brandName[bid]; });
    }
    const sYq = sYs.map((s) => s.q);
    let res = compute(o, meta, P, sCur.q, sYq, brandOf);
    if (o.boxRound && res.rows.length) {
      const pids = [...new Set(res.rows.map((r) => r.pid).filter(Boolean))];
      msg(`Διαβάζω «Ποσότητα» κιβωτίου για ${pids.length} είδη…`);
      const packOf = await packQuantities(pids, (d, t) => msg(`Διαβάζω «Ποσότητα» κιβωτίου… ${d} / ${t}`));
      res = compute(o, meta, P, sCur.q, sYq, brandOf, packOf);
      res.packCount = Object.keys(packOf).length;
    }
    const wb = await buildExcel(o, meta, res, periods);
    const buf = await wb.xlsx.writeBuffer();
    const short = meta.active.name.replace(/^NIKOLAS NAOUSA\s*/i, '').replace(/[^\wΑ-Ωα-ω]+/g, '_') || 'STORE';
    const vtag = (o.vendor ? '_' + (meta.vendors[o.vendor] || o.vendor).replace(/[^\wΑ-Ωα-ω]+/g, '_').slice(0, 20) : '') +
                 (o.brand ? '_' + (meta.brandName[o.brand] || o.brand).replace(/[^\wΑ-Ωα-ω]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 20) : '');
    const fname = `Paraggelia_${short}${vtag}_${o.from.replace(/-/g, '')}-${o.to.replace(/-/g, '')}${o.cheapest ? '_fthinoteroi' : ''}.xlsx`;
    if (download) {
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fname; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    }
    return { res, fname, bytes: buf.byteLength, periods, buf };
  }

  // Αυτόματη αποστολή e-mail: Google Apps Script στον λογαριασμό του χρήστη (στέλνει μόνο σε επιτρεπτούς παραλήπτες).
  // Ο σύνδεσμος και το κλειδί αποθηκεύονται ΜΟΝΟ σε αυτόν τον browser (όχι στον δημόσιο κώδικα).
  const hookUrl = () => { const u = store.get('mailHook', ''), k = store.get('mailKey', ''); return u && k ? u + (u.includes('?') ? '&' : '?') + 'k=' + encodeURIComponent(k) : ''; };
  const hookSt = () => { const on = !!hookUrl(); $id('spo-hook').value = store.get('mailHook', '');
    $id('spo-hook-st').innerHTML = on ? '<b style="color:#2e7d32">✔ Ενεργή</b> — τα κουμπιά «✉️» στέλνουν την παραγγελία αμέσως (με το Excel συνημμένο), από rrms2021@gmail.com.'
      : (store.get('mailKey', '') ? 'Ανενεργή — επικολλήστε το «Web app URL» του script «Παραγγελίες e-mail» και πατήστε OK. Μέχρι τότε ανοίγει έτοιμο μήνυμα στο Gmail.'
        : 'Ανενεργή — λείπει το κλειδί σε αυτόν τον browser. Μέχρι τότε ανοίγει έτοιμο μήνυμα στο Gmail.'); };
  hookSt();
  $id('spo-hook-save').addEventListener('click', async () => {
    let v = $id('spo-hook').value.trim();
    if (!v) { store.set('mailHook', ''); hookSt(); return; }
    const km = v.match(/[?&]k=([^&#]+)/); if (km) { store.set('mailKey', decodeURIComponent(km[1])); }
    v = v.replace(/[?#].*$/, '');
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(v)) { $id('spo-hook-st').innerHTML = '<b style="color:#b71c1c">Ο σύνδεσμος πρέπει να είναι της μορφής https://script.google.com/macros/s/…/exec</b>'; return; }
    store.set('mailHook', v); $id('spo-hook-st').textContent = 'Έλεγχος σύνδεσης…';
    try { const j = await fetch(hookUrl()).then((r) => r.json());
      if (!j.ready) throw new Error('δεν απάντησε σωστά'); if (!j.keyOk) throw new Error('λάθος κλειδί');
      hookSt(); $id('spo-hook-st').insertAdjacentHTML('beforeend', ' (έλεγχος OK — δεν στάλθηκε τίποτα)');
    } catch (e) { $id('spo-hook-st').innerHTML = '<b style="color:#b71c1c">Η σύνδεση απέτυχε: ' + esc(e.message || e) + '.</b> Ελέγξτε ότι η δημοσίευση είναι «Web app», «Execute as: Me», «Who has access: Anyone».'; }
  });

  // Γρήγορα κουμπιά: τελευταίες 7 ημέρες, ο προμηθευτής, «Φθηνότερος = ΟΧΙ», όλοι οι κατασκευαστές, καταχώρηση στο site
  panel.querySelectorAll('.qb button').forEach((b) => b.addEventListener('click', () => {
    const q = b.dataset.q, t = new Date(); t.setDate(t.getDate() - 1); const f = new Date(t); f.setDate(f.getDate() - 6);
    $id('spo-from').value = isoOf(f); $id('spo-to').value = isoOf(t);
    $id('spo-brand').value = '';
    const sd = SEND[q];
    $id('spo-vendor').value = sd ? sd.vid : '';
    $id('spo-send').value = sd ? q : '';
    if (sd) { const sg = panel.querySelector('.seg[data-key="cheapest"]'); [...sg.children].forEach((x) => x.classList.toggle('on', x.dataset.v === '0')); }
    // το παράθυρο ανοίγει ΤΩΡΑ (με το πάτημα), αλλιώς ο browser το μπλοκάρει· η διεύθυνση μπαίνει όταν είναι έτοιμη η παραγγελία
    let win = null; if (sd && !(sd.email && hookUrl())) { try { win = window.open('about:blank', '_blank'); if (win) win.document.write('<p style="font:16px Arial;padding:20px">Ετοιμάζεται η παραγγελία από το SimplyPOS…</p>'); } catch (e) { win = null; } }
    doRun(win);
  }));
  $id('spo-go').addEventListener('click', () => doRun(null));
  async function doRun(win) {
    const o = opts(); store.set('opts', o);
    const btn = $id('spo-go'); btn.disabled = true; btn.textContent = 'Υπολογισμός…'; msg('Διαβάζω πωλήσεις και αποθέματα…');
    panel.querySelectorAll('.qb button').forEach((b) => { b.disabled = true; });
    try {
      const { res, fname, buf } = await run(o, true);
      const sum = (k) => res.rows.reduce((a, r) => a + k(r), 0);
      if (!res.rows.length) { msg('Δεν βγήκε καμία γραμμή παραγγελίας με αυτές τις επιλογές. Το Excel κατέβηκε κενό.'); }
      else msg(`<div class="big">✔ ${esc(fname)}</div>
        <div>${res.rows.length} είδη · ${fmtInt(sum((r) => r.q))} τεμ. · ${res.vorder.length} προμηθευτές</div>
        <div>Καθαρή αξία <b>${fmtEur(sum((r) => r.val))}</b> · με ΦΠΑ <b>${fmtEur(sum((r) => r.gross))}</b></div>
        <table>${res.vorder.slice(0, 12).map((v) => `<tr><td>${esc(v)}</td><td class="r">${fmtEur(res.tot[v])}</td></tr>`).join('')}
        ${res.vorder.length > 12 ? `<tr><td colspan="2">… και ${res.vorder.length - 12} ακόμη</td></tr>` : ''}</table>`);
      const sd = SEND[o.send];
      if (sd && sd.email && res.rows.length) {
        // e-mail: έτοιμο μήνυμα στο Gmail (δεν στέλνεται αυτόματα — ο χρήστης πατάει «Αποστολή»)
        const now = new Date(), dstr = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
        const subject = `Παραγγελία ${meta.active.name} — ${dstr}`;
        const lines = res.rows.map((r, i) => `${i + 1}. ${r.nm} — ${fmtInt(r.q)} τεμ.` +
          (r.boxes ? ` (${r.boxes} κιβ. × ${r.pk})` : '') + (r.code ? `   [κωδ. ${r.code}]` : ''));
        const body = `Καλησπέρα σας,\n\nπαρακαλώ την παρακάτω παραγγελία για το κατάστημα ${meta.active.name}:\n\n` +
          lines.join('\n') + `\n\nΣύνολο: ${res.rows.length} είδη, ${fmtInt(sum((r) => r.q))} τεμάχια.\n\nΕυχαριστώ,\n${meta.active.name}`;
        const enc = encodeURIComponent;
        // authuser = λογαριασμός Gmail αποστολέα (ανοίγει αυτός, ακόμη κι αν στον Chrome υπάρχουν κι άλλοι)
        const GM = `https://mail.google.com/mail/?authuser=${enc(sd.from || '')}&view=cm&fs=1&to=`;
        let gm = `${GM}${enc(sd.email)}&su=${enc(subject)}&body=${enc(body)}`;
        if (gm.length > 7500) gm = `${GM}${enc(sd.email)}&su=${enc(subject)}&body=${enc('Καλησπέρα σας,\n\nπαρακαλώ την παραγγελία στο συνημμένο Excel (' + res.rows.length + ' είδη).\n\nΕυχαριστώ,\n' + meta.active.name)}`;
        const mt = `mailto:${sd.email}?subject=${enc(subject)}&body=${enc(body.length > 1800 ? 'Η παραγγελία είναι στο συνημμένο Excel.' : body)}`;
        let sentOk = false;
        if (hookUrl()) {
          $id('spo-msg').insertAdjacentHTML('beforeend', '<div id="spo-sending" style="margin-top:10px;padding:10px;border-radius:8px;background:#fff8d6">✉️ Αποστολή e-mail στον ' + esc(sd.label) + '…</div>');
          try {
            const u8 = new Uint8Array(buf); let bin = ''; for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
            const r = await fetch(hookUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
              body: JSON.stringify({ to: sd.email, subject, body, xlsx: btoa(bin), filename: fname }) });
            const j = await r.json();
            if (!j.ok) throw new Error(j.error || 'άγνωστο σφάλμα');
            sentOk = true;
            $id('spo-sending').outerHTML = `<div style="margin-top:10px;padding:12px;border-radius:8px;background:#e8f5e9;font-size:14px">
              <b>✔ Η παραγγελία στάλθηκε</b> στον ${esc(sd.label)} (${esc(sd.email)}) από ${esc(sd.from || '')} — ${res.rows.length} είδη, με το Excel συνημμένο.
              <div class="hint">Θα τη βρείτε στα «Απεσταλμένα» του Gmail.</div></div>`;
          } catch (e) {
            $id('spo-sending').outerHTML = `<div style="margin-top:10px;padding:10px;border-radius:8px;background:#fde2e2"><b>Η αυτόματη αποστολή απέτυχε:</b> ${esc(e.message || e)}. Στείλτε την από το Gmail με το παρακάτω κουμπί.</div>`;
          }
        }
        if (!sentOk) {
        $id('spo-msg').insertAdjacentHTML('beforeend', `<div style="margin-top:10px;padding:10px;border-radius:8px;background:#e8f5e9">
          <b>E-mail προς ${esc(sd.label)}</b> (${esc(sd.email)}) από ${esc(sd.from || '')}: ${res.rows.length} είδη.<br>
          <a href="${gm}" target="_blank" rel="noopener" style="display:inline-block;margin-top:6px;padding:8px 12px;background:#2e7d32;color:#fff;border-radius:8px;font-weight:bold;text-decoration:none">Άνοιγμα στο Gmail ›</a>
          <a href="${mt}" style="display:inline-block;margin:6px 0 0 6px;padding:8px 12px;background:#1F4E78;color:#fff;border-radius:8px;font-weight:bold;text-decoration:none">Πρόγραμμα e-mail</a>
          <button type="button" id="spo-copymail" style="margin:6px 0 0 6px;padding:8px 12px;border:1px solid #c5ced8;border-radius:8px;background:#fff;cursor:pointer;font-weight:bold">Αντιγραφή κειμένου</button>
          <div class="hint">Το μήνυμα ΔΕΝ στέλνεται αυτόματα: ελέγξτε το και πατήστε «Αποστολή». Αν θέλετε, επισυνάψτε το Excel που κατέβηκε (${esc(fname)}).</div></div>`);
        $id('spo-copymail').onclick = () => { const t = subject + '\n\n' + body;
          (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(() => { $id('spo-copymail').textContent = '✔ Αντιγράφηκε'; },
            () => { prompt('Αντιγραφή (Ctrl+C):', t); }); };
        if (win && !win.closed) { try { win.location.href = gm; } catch (e) {} }
        } else if (win && !win.closed) win.close();
      } else if (sd && !sd.email && res.rows.length) {
        // η παραγγελία πάει στο site μέσα στο #hash του συνδέσμου (δεν φεύγει προς κανέναν server)
        const payload = { v: 2, to: o.send, store: meta.active.name, period: `${isoToGr(o.from)}–${isoToGr(o.to)}`,
          rows: res.rows.map((r) => [r.cd, r.nm, r.q, r.pk || 0, r.boxes == null ? 0 : r.boxes,
            ((r.scl || []).find((c) => c[0] === sd.vid) || [])[1] || '', (r.bcs || []).join(',')]) };
        // base64url: μόνο γράμματα/αριθμοί/-/_ ώστε κανένα site να μην «χαλάει» τον σύνδεσμο
        const b64u = btoa(unescape(encodeURIComponent(JSON.stringify(payload)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
        const url = sd.url + '#spo=' + b64u;
        $id('spo-msg').insertAdjacentHTML('beforeend', `<div style="margin-top:10px;padding:10px;border-radius:8px;background:#e8f5e9">
          <b>Καταχώρηση στον/στην ${esc(sd.label)}:</b> ${res.rows.length} είδη.<br>
          <a href="${url}" target="_blank" rel="noopener" style="display:inline-block;margin-top:6px;padding:8px 12px;background:#2e7d32;color:#fff;border-radius:8px;font-weight:bold;text-decoration:none">Άνοιγμα ${esc(sd.label)} ›</a>
          <div class="hint">Στο site που ανοίγει πατήστε ξανά τον σελιδοδείκτη «📦 Παραγγελίες».</div></div>`);
        if (win && !win.closed) { try { win.location.href = url; } catch (e) {} }
      } else if (win && !win.closed) win.close();

    } catch (e) { msg(esc(e.message || e), true); if (win && !win.closed) win.close(); }
    btn.disabled = false; btn.textContent = 'ΔΗΜΙΟΥΡΓΙΑ EXCEL';
    panel.querySelectorAll('.qb button').forEach((b) => { b.disabled = false; });
  }

  window.__spOrders = { run, compute, opts, meta };


  // =====================================================================
  //  Site προμηθευτή (ΡΑΓΚΟΥΣΗΣ / Mmmgroup / ΜΑΣΟΥΤΗΣ): η παραγγελία μπαίνει ΜΟΝΟ στο καλάθι.
  //  Την αποστολή της παραγγελίας την κάνει πάντα ο χρήστης από το site.
  // =====================================================================
  async function supplierModule(which) {
    const RID = 'spo-sup-panel';
    const old = document.getElementById(RID); if (old) { old.style.display = 'flex'; return; }
    const $ = (x) => document.getElementById(x);
    const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const ls = { get(k, d) { try { const v = localStorage.getItem('spo.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
                 set(k, v) { try { localStorage.setItem('spo.' + k, JSON.stringify(v)); } catch (e) {} } };
    const unq = (t) => (t && t.startsWith('"') ? JSON.parse(t) : t);
    const packOfName = (n) => { const m = String(n).match(/(\d+)\s*[*xXΧχ×]\s*(\d+)\s*(ml|lt|l|gr|g|cl|cc)?/i); if (!m) return 0; return m[3] ? +m[1] : (+m[1]) * (+m[2]); };

    // ---------------- προσαρμογή ανά site ----------------
    // ΜΑΣΟΥΤΗΣ Cash&Carry (expressapp.masoutis.gr): χωρίς πλήρη κατάλογο — αναζήτηση ανά είδος (barcode / κωδικός / όνομα)
    const MAS = {
      title: 'ΜΑΣΟΥΤΗ',
      builtin: {},
      loggedIn: () => !!localStorage.getItem('token'),
      _c: null, _H: null,
      async creds() {
        if (this._c) return;
        const c = await fetch('/api/auth/GetCreds/', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + localStorage.getItem('token') }, body: '{}' }).then((r) => r.json());
        if (!c || !c.IfServiceOk) throw new Error('Δεν έγινε σύνδεση με το site του Μασούτη. Συνδεθείτε ξανά και πατήστε ξανά τον σελιδοδείκτη.');
        this._c = c; this._H = { 'Content-Type': 'application/json', Key: c.Key, Uid: c.Uid, Usl: c.Usl };
      },
      async post(path, body) { await this.creds(); const r = await fetch('/api/WebOrders/' + path, { method: 'POST', headers: this._H, body: JSON.stringify(body) }); if (!r.ok) throw new Error('Σφάλμα ' + r.status + ' (' + path + ')'); return r.json(); },
      map(x) { const bcs = {}; if (x.Barcode) bcs[String(x.Barcode)] = 1;
        return { key: x.Itemcode, sku: String(x.Itemcode), name: String(x.ItemDescr || ''), unit: String(x.MI1MME || 'τεμ').toUpperCase(), price: Number(x.UnitPrice) || 0,
          inStock: true, bcs, raw: x }; },
      async search(q) { const j = await this.post('GetItemList/', { CustomerToken: this._c ? this._c.Token : (await this.creds(), this._c.Token), SortId: 0, IfHomePage: false, IfPromo: 0, Category: '', SearchText: q });
        return (j.ItemList || []).map((x) => this.map(x)); },
      async load(rows, progress) {
        await this.creds();
        const found = {}; let i = 0, done = 0;
        const words = (n) => String(n || '').replace(/[^0-9A-Za-zΑ-Ωα-ωΆ-ώ ]+/g, ' ').split(/\s+/).filter((w) => w.length >= 3 && !/^\d+$/.test(w)).slice(0, 3).join(' ');
        const worker = async () => {
          while (i < rows.length) {
            const a = rows[i++]; const bcs = String(a[6] || '').split(',').filter(Boolean), rk = String(a[5] || '');
            let hit = false;
            for (const q of [...new Set([...bcs, rk].filter(Boolean))]) {
              const l = await this.search(q).catch(() => []);
              l.forEach((p) => { found[p.sku] = p; });
              if (l.some((p) => bcs.includes(Object.keys(p.bcs)[0]) || p.sku === rk)) { hit = true; break; }
            }
            if (!hit && words(a[1])) (await this.search(words(a[1])).catch(() => [])).slice(0, 12).forEach((p) => { found[p.sku] = p; });
            done++; if (progress) progress(done, rows.length);
          }
        };
        await Promise.all(Array.from({ length: 6 }, worker));
        return Object.values(found);
      },
      qty(row, p, bcMul) {
        const x = p.raw || {}, step = Number(x.OrderQty) || 1, max = Number(x.MaxQty) || 9999;
        let q = (Number(row.q) || 0) * (bcMul || 1);
        // Το «τεμ» του Μασούτη μπορεί να είναι συσκευασία (π.χ. «12X500ML»). Αν το barcode δεν είναι το ίδιο με του SimplyPOS
        // και το όνομα στο site δείχνει συσκευασία που δεν έχει το είδος στο SimplyPOS → τεμάχια / συσκευασία.
        const sameBc = (row.bcs || []).includes(String(x.Barcode || ''));
        const pkSite = packOfName(p.name), pkPos = packOfName(row.nm);
        if (!sameBc && pkSite > 1 && pkSite !== pkPos) q = Math.max(1, Math.round(q / pkSite));
        return Math.min(max, Math.max(step, Math.round(q / step) * step));
      },
      async addToCart(items, empty) {
        const T = () => this._c.Token;
        if (empty) await this.post('ClearBasket', { CustomerToken: T(), CustomerId: '', CustomerPin: '', CustomerMobile: '', IfMobile: false });
        const cur = {}; if (!empty) { const b = await this.post('SyncBasket/', { CustomerToken: T() }); (b.Details || []).forEach((d) => { cur[d.Itemcode] = Number(d.UserQty) || 0; }); }
        const fails = [];
        for (const { p, qty } of items) {
          const x = p.raw, q = (cur[x.Itemcode] || 0) + qty;
          const it = { Itemcode: x.Itemcode, UserQty: q, PosPrice: Math.round(x.UnitPrice * q * 100) / 100, OrderQty: x.OrderQty, ItemDescr: x.ItemDescr, Image: x.Image, UnitPrice: x.UnitPrice };
          const u = await this.post('UpdateBasket/', { ServiceResponse: T(), Item: it }).catch((e) => ({ IfServiceOk: false, ServiceResponse: e.message }));
          if (!u.IfServiceOk) fails.push(p.name + (u.ServiceResponse ? ' (' + u.ServiceResponse + ')' : ''));
        }
        const b = await this.post('SyncBasket/', { CustomerToken: T() });
        return { lines: Number(b.BasketCount) || (b.Details || []).length, total: Number(b.BasketAmount) || 0, net: false, fails, reload: true };
      }
    };
    const A = which === 'mas' ? MAS : which === 'mmm' ? {
      title: 'Mmmgroup',
      // Αντιστοιχίσεις που έχουν επιβεβαιωθεί (Κωδ. SimplyPOS → κωδικός είδους στο site)
      builtin: {},
      loggedIn: () => !!localStorage.getItem('api_bearer_token'),
      async load() {
        let anon = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR0dmdoa2ZmeHlvYmlqYXZtZWZ0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njk3NTc1OTksImV4cCI6MjA4NTMzMzU5OX0.oG8ug6M-zpxcloaODnSZ7rkW8UrGdESoXIvBO_RMGCw';
        try { const src = [...document.scripts].map((s) => s.src).find((s) => /\/assets\/index-/.test(s));
          if (src) { const m = (await fetch(src).then((r) => r.text())).match(/eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9\.[\w-]+\.[\w-]+/); if (m) anon = m[0]; } } catch (e) {}
        const tok = unq(localStorage.getItem('api_bearer_token'));
        const r = await fetch('https://dtvghkffxyobijavmeft.supabase.co/functions/v1/catalog?scope=main',
          { headers: { apikey: anon, Authorization: 'Bearer ' + (tok || anon), Accept: 'application/json' } });
        if (!r.ok) throw new Error('Σφάλμα ' + r.status + ' στον κατάλογο της Mmmgroup. Συνδεθείτε ξανά.');
        const j = await r.json();
        return (j.products || []).map((p) => {
          const bcs = {}; if (p.barcode) bcs[p.barcode] = 1;
          try { JSON.parse(p.multiple_barcodes || '[]').forEach((m) => { if (m && m.barcode && !(m.barcode in bcs)) bcs[m.barcode] = Number(m.quantity) || 1; }); } catch (e) {}
          return { key: String(p.id).replace(/^prd/, ''), sku: String(p.sku || ''), name: String(p.title || ''), unit: 'ΤΕΜ', price: Number(p.final_price ?? p.base_price) || 0,
            inStock: Number(p.total_stock) > 0, bcs, raw: p };
        });
      },
      qty(row, p, bcMul) { return Math.max(1, Math.round((Number(row.q) || 0) * (bcMul || 1))); },
      async addToCart(items, empty) {
        // Το καλάθι της Mmmgroup ζει στον browser (localStorage «mmm_cart»)· μετά γίνεται ανανέωση της σελίδας.
        let cart = []; if (!empty) { try { cart = JSON.parse(localStorage.getItem('mmm_cart') || '[]') || []; } catch (e) { cart = []; } }
        items.forEach(({ p, qty }) => {
          const x = p.raw, price = Number(x.final_price ?? x.base_price) || 0, base = Number(x.base_price) || undefined;
          const prod = { id: p.key, name: p.name, category: '', categorySlug: '', price, originalPrice: base && base > price + 0.005 ? base : undefined,
            netPrice: price, netOriginalPrice: base && base > price + 0.005 ? base : undefined, imageUrl: x.image_path || '', images: x.image_path ? [x.image_path] : [],
            isNew: false, isPopular: false, madeToOrder: false, available: p.inStock, isOnSale: !!(base && base > price + 0.005), stock: Number(x.total_stock) || 0,
            stockKnown: true, sku: p.sku, barcode: x.barcode || '', brand: x.manufacturerbrand_name || '', slug: x.slug || '', vatRate: Number(x.rate) || 0,
            createdAt: x.created_at || '' };
          const ex = cart.find((c) => c && c.product && String(c.product.sku) === p.sku);
          if (ex) ex.quantity = (Number(ex.quantity) || 0) + qty; else cart.push({ product: prod, quantity: qty });
        });
        localStorage.setItem('mmm_cart', JSON.stringify(cart));
        const tot = cart.reduce((a, c) => a + (Number(c.product.price) || 0) * (Number(c.quantity) || 0), 0);
        return { lines: cart.length, total: tot, net: true, fails: [], reload: true };
      }
    } : {
      title: 'ΡΑΓΚΟΥΣΗ',
      builtin: { '6880': '02.172.0078.0005', '1333': '02.172.0078.0011', '6451': '02.172.0078.0006', '1698': '02.135.0078.0014',
        '4044': '02.135.0078.0007', '1695': '02.135.0078.0012', '9431': '02.107.0078.0037', '1320': '02.168.0078.0003', '1945': '02.135.0078.0017',
        '9319': '02.135.0078.0035', '3396': '02.135.0078.0006', '2568': '02.172.0050.0013', '2556': '02.107.0050.0002', '8030': '02.113.0050.0015', '3051': '02.113.0050.0011' },
      loggedIn: () => !!localStorage.getItem('b2b_token'),
      H() { return { Authorization: 'Bearer ' + unq(localStorage.getItem('b2b_token')), 'Content-Type': 'application/json' }; },
      async load() {
        const out = [];
        for (let pg = 1, tp = 1; pg <= tp && pg < 80; pg++) {
          const r = await fetch(`/api/products?limit=500&page=${pg}`, { headers: this.H() });
          if (!r.ok) throw new Error('Σφάλμα ' + r.status + ' στη λίστα ειδών. Συνδεθείτε ξανά στο site.');
          const j = await r.json(); tp = j.totalPages || 1;
          (j.products || []).forEach((p) => out.push({ key: p.id, sku: p.sku, name: p.name, unit: p.unit, price: Number(p.unitPrice) || 0, inStock: !!p.inStock, bcs: {}, raw: p }));
        }
        return out;
      },
      qty(row, p) {
        const q = Number(row.q) || 0;
        if (p.unit === 'ΤΕΜ') return Math.max(1, Math.round(q));
        if (/6ΑΔΑ|ΕΞΑΔΑ/i.test(row.nm)) return Math.max(1, Math.round(q));   // πωλείται ήδη ως συσκευασία
        const pk = packOfName(p.name) || Number(row.pk) || 0;
        return pk > 0 ? Math.max(1, Math.round(q / pk)) : Math.max(1, Math.round(q));
      },
      async addToCart(items, empty) {
        const H = this.H();
        if (empty) { const cur = await fetch('/api/cart', { headers: H }).then((r) => r.json()); for (const it of (cur.items || [])) await fetch('/api/cart/items/' + it.productId, { method: 'DELETE', headers: H }); }
        const fails = [];
        for (const { p, qty } of items) {
          const r = await fetch('/api/cart/items', { method: 'POST', headers: H, body: JSON.stringify({ productId: p.key, quantity: qty }) });
          if (!r.ok) fails.push(p.name);
        }
        const c = await fetch('/api/cart', { headers: H }).then((r) => r.json());
        return { lines: (c.items || []).length, total: Number(c.grandTotal) || 0, net: false, fails, reload: false };
      }
    };

    // ---------------- παραγγελία από το κουμπί του SimplyPOS ----------------
    let order = null, fresh = false;
    const hm = (location.hash + ' ' + location.href).match(/spo=([A-Za-z0-9_%\-]+)/);
    const dec = (t) => { try { let b = t.replace(/-/g, '+').replace(/_/g, '/'); while (b.length % 4) b += '='; return JSON.parse(decodeURIComponent(escape(atob(b)))); }
      catch (e) { return JSON.parse(decodeURIComponent(t)); } };
    if (hm) { try { order = dec(hm[1]); fresh = true; ls.set('lastOrder', order); history.replaceState(null, '', location.pathname + location.search); } catch (e) { order = null; } }
    if (!order) order = ls.get('lastOrder', null);
    const learned = ls.get('map', {});

    const css = `#${RID}{position:fixed;top:12px;right:12px;bottom:12px;z-index:2147483646;width:min(980px,calc(100vw - 24px));display:flex;flex-direction:column;
      background:#fff;color:#1f2933;border-radius:14px;box-shadow:0 12px 40px rgba(0,0,0,.45);font:13px/1.35 Arial,Helvetica,sans-serif;overflow:hidden}
      #${RID} *{box-sizing:border-box;font-family:inherit;color:inherit}
      #${RID} .hd{background:#1F4E78;color:#fff;padding:10px 14px;display:flex;justify-content:space-between;align-items:center}
      #${RID} .hd *{color:#fff} #${RID} .hd button{background:none;border:0;font-size:20px;cursor:pointer}
      #${RID} .bd{padding:10px 14px;overflow:auto;flex:1}
      #${RID} .ft{padding:10px 14px;border-top:1px solid #e4e7eb;display:flex;gap:10px;align-items:center;flex-wrap:wrap;background:#f8fafc}
      #${RID} table{width:100%;border-collapse:collapse} #${RID} th{position:sticky;top:-10px;background:#1F4E78;color:#fff;padding:6px 4px;text-align:left;font-size:12px}
      #${RID} td{border-bottom:1px solid #e4e7eb;padding:4px;vertical-align:middle}
      #${RID} tr.ok td{background:#fff} #${RID} tr.chk td{background:#fff8d6} #${RID} tr.none td{background:#fde2e2} #${RID} tr.off td{opacity:.45}
      #${RID} select,#${RID} input[type=number],#${RID} input[type=text]{border:1px solid #c5ced8;border-radius:6px;padding:4px;font-size:12px;background:#fff;color:#1f2933}
      #${RID} select{width:100%} #${RID} input[type=number]{width:64px} #${RID} input.sr{width:100%;margin-top:3px}
      #${RID} .go{padding:10px 16px;border:0;border-radius:10px;background:#2e7d32;color:#fff;font-weight:bold;font-size:14px;cursor:pointer}
      #${RID} .go:disabled{background:#9aa5b1} #${RID} .tag{font-size:11px;border-radius:4px;padding:1px 5px;background:#eef2f6;white-space:nowrap}
      #${RID} .warn{color:#b71c1c;font-weight:bold}`;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    const panel = document.createElement('div'); panel.id = RID;
    panel.innerHTML = `<div class="hd"><b>📦 Παραγγελία → καλάθι ${esc(A.title)}</b><button data-x="1" title="Κλείσιμο">×</button></div>
      <div class="bd" id="spo-s-bd">Φόρτωση…</div>
      <div class="ft"><label><input type="checkbox" id="spo-s-empty" checked> Άδειασμα καλαθιού πριν</label>
        <button class="go" id="spo-s-go" disabled>ΠΡΟΣΘΗΚΗ ΣΤΟ ΚΑΛΑΘΙ</button><span id="spo-s-msg"></span>
        <span style="margin-left:auto;font-size:11px;color:#616e7c">Η παραγγελία ΔΕΝ στέλνεται — την στέλνετε εσείς από το καλάθι του site.</span></div>`;
    document.body.appendChild(panel);
    panel.querySelector('[data-x]').onclick = () => { panel.style.display = 'none'; };
    const bd = $('spo-s-bd'), smsg = (h) => { $('spo-s-msg').innerHTML = h; };

    if (!A.loggedIn()) { bd.innerHTML = `<p class="warn">Συνδεθείτε πρώτα στο site (${esc(A.title)}) και πατήστε ξανά τον σελιδοδείκτη.</p>`; return; }
    if (!order || !order.rows || !order.rows.length) {
      bd.innerHTML = `<p>Δεν βρέθηκε παραγγελία. Στο SimplyPOS ανοίξτε το κουμπί «📦 Παραγγελίες», στο πεδίο 9 διαλέξτε <b>${esc(A.title)}</b>,
        πατήστε «ΔΗΜΙΟΥΡΓΙΑ EXCEL» και μετά το πράσινο κουμπί «Άνοιγμα …».</p>`;
      return;
    }
    if (order.to && order.to !== which) {
      bd.innerHTML = `<p class="warn">Η παραγγελία που έφτιαξε το SimplyPOS είναι για άλλον προμηθευτή. Φτιάξτε την ξανά διαλέγοντας ${esc(A.title)} στο πεδίο 9.</p>`;
      return;
    }

    let prods;
    try { bd.textContent = 'Φόρτωση ειδών του site…'; prods = await A.load(order.rows, (d, t) => { bd.textContent = `Αναζήτηση ειδών στο site… ${d} / ${t}`; }); }
    catch (e) { bd.innerHTML = '<p class="warn">' + esc(e.message || e) + '</p>'; return; }
    const bySku = {}, byKey = {}, byDig = {}, byBc = {}; const dig = (s) => String(s || '').replace(/\D/g, '');
    const indexP = (p) => { bySku[p.sku] = p; byKey[String(p.key)] = p; const d = dig(p.sku); if (d) (byDig[d] = byDig[d] || []).push(p);
      Object.keys(p.bcs).forEach((b) => { if (!byBc[b]) byBc[b] = [p, p.bcs[b]]; }); };
    prods.forEach(indexP);

    // ---------------- ταίριασμα ονομάτων ----------------
    const LAT = { 'Α': 'A', 'Β': 'B', 'Ε': 'E', 'Ζ': 'Z', 'Η': 'H', 'Ι': 'I', 'Κ': 'K', 'Μ': 'M', 'Ν': 'N', 'Ο': 'O', 'Ρ': 'P', 'Τ': 'T', 'Υ': 'Y', 'Χ': 'X' };
    const norm = (s) => String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[ΑΒΕΖΗΙΚΜΝΟΡΤΥΧ]/g, (c) => LAT[c]);
    const vol = (s) => { const m = norm(s).match(/(\d+(?:[.,]\d+)?)\s*(ML|MΛ|CC|LT|LIT|L|ΛΤ)\b/); if (!m) return 0; const v = parseFloat(m[1].replace(',', '.')); return /ML|MΛ|CC/.test(m[2]) ? v : v * 1000; };
    const STOP = new Set(['ML', 'LT', 'PET', 'CAN', 'NC', 'TEM', 'KIB', 'X', 'THE']);
    const toks = (s) => new Set(norm(s).replace(/(\d+(?:[.,]\d+)?)\s*(ML|MΛ|CC|LT|LIT|L|ΛΤ)\b/g, ' ').split(/[^A-Z0-9Α-Ω]+/).filter((t) => t.length >= 2 && !STOP.has(t) && !/^\d+$/.test(t)));
    const df = {}; prods.forEach((p) => { p._t = toks(p.name); p._v = vol(p.name); p._t.forEach((t) => { df[t] = (df[t] || 0) + 1; }); });
    const idf = (t) => Math.log((prods.length + 1) / ((df[t] || 0) + 1));   // σπάνιες λέξεις (μάρκα) μετράνε περισσότερο
    const score = (nm, p) => {
      const a = toks(nm); if (!a.size) return 0; let wi = 0, wa = 0, inter = 0;
      a.forEach((t) => { const w = idf(t); wa += w; if (p._t.has(t)) { wi += w; inter++; } });
      let s = 0.8 * wi / wa + 0.2 * inter / (a.size + p._t.size - inter);
      const v = vol(nm); if (v && p._v) s += (Math.abs(v - p._v) < 1 ? 0.2 : -0.3);
      return s;
    };
    const suggest = (q, n) => prods.map((p) => [score(q, p), p]).filter((x) => x[0] > 0.15).sort((a, b) => b[0] - a[0]).slice(0, n || 8);

    // ---------------- γραμμές ----------------
    const R = order.rows.map((a) => ({ cd: String(a[0]), nm: a[1], q: a[2], pk: a[3], bx: a[4], rk: a[5] || '', bcs: String(a[6] || '').split(',').filter(Boolean) }));
    R.forEach((r) => {
      let p = null, how = '', mul = 1;
      if (learned[r.cd] && bySku[learned[r.cd]]) { p = bySku[learned[r.cd]]; how = 'αποθηκευμένη αντιστοίχιση'; }
      else if (A.builtin[r.cd] && bySku[A.builtin[r.cd]]) { p = bySku[A.builtin[r.cd]]; how = 'αντιστοίχιση'; }
      if (!p) {   // barcode: προτιμάται το barcode τεμαχίου (×1) από barcode κιβωτίου
        const hits = r.bcs.map((b) => byBc[b] || byBc[b.replace(/^0+/, '')]).filter(Boolean).sort((x, y) => (x[1] || 1) - (y[1] || 1));
        if (hits.length) { p = hits[0][0]; mul = hits[0][1] || 1; how = 'barcode' + (mul > 1 ? ` (κιβώτιο ×${mul})` : ''); }
      }
      if (!p && r.rk && bySku[r.rk]) { p = bySku[r.rk]; how = 'κωδικός'; }
      if (!p && r.rk && dig(r.rk).length >= 8 && (byDig[dig(r.rk)] || []).length === 1) { p = byDig[dig(r.rk)][0]; how = 'κωδικός'; }
      r.mul = mul;
      r.sug = p ? [] : suggest(r.nm);
      if (p) { r.sel = p.sku; r.st = 'ok'; }
      else if (r.sug.length && r.sug[0][0] >= 0.9) { r.sel = r.sug[0][1].sku; r.st = 'chk'; how = 'πρόταση από όνομα — ελέγξτε'; }
      else { r.sel = ''; r.st = 'none'; how = 'δεν βρέθηκε — διαλέξτε'; }
      r.how = how; r.on = true; r.qty = r.sel ? A.qty(r, bySku[r.sel], r.mul) : 0;
    });

    const optHtml = (r) => {
      const list = []; const seen = new Set();
      const add = (p, sc) => { if (!p || seen.has(p.sku)) return; seen.add(p.sku); list.push(`<option value="${esc(p.sku)}"${p.sku === r.sel ? ' selected' : ''}>${esc(p.name)} · ${esc(p.unit)} · ${p.price.toFixed(2)}€${p.inStock ? '' : ' · ΕΞΑΝΤΛ.'}${sc != null ? ` (${Math.round(sc * 100)}%)` : ''}</option>`); };
      if (r.sel) add(bySku[r.sel]); r.sug.forEach(([sc, p]) => add(p, sc));
      return `<option value=""${r.sel ? '' : ' selected'}>— να μην μπει —</option>` + list.join('');
    };
    const render = () => {
      const c = { ok: 0, chk: 0, none: 0 }; R.forEach((r) => { c[r.st]++; });
      bd.innerHTML = `<div style="margin-bottom:8px">Παραγγελία από SimplyPOS: <b>${esc(order.store || '')}</b> · πωλήσεις ${esc(order.period || '')} · <b>${R.length}</b> είδη —
        <span class="tag">✔ ${c.ok} βρέθηκαν</span> <span class="tag" style="background:#fff3b0">⚠ ${c.chk} προτάσεις από όνομα</span>
        <span class="tag" style="background:#fde2e2">✖ ${c.none} δεν βρέθηκαν</span>. Ελέγξτε τις κίτρινες/κόκκινες γραμμές· οι επιλογές σας αποθηκεύονται για την επόμενη φορά.</div>
        <table><thead><tr><th style="width:28px"></th><th>Είδος SimplyPOS</th><th>Τεμ.</th><th style="width:46%">Είδος στο site ${esc(A.title)}</th><th>Ποσότητα στο site</th></tr></thead><tbody>
        ${R.map((r, i) => {
          const p = bySku[r.sel];
          return `<tr class="${r.on ? r.st : 'off'}" data-i="${i}"><td><input type="checkbox" class="on" ${r.on ? 'checked' : ''}></td>
            <td>${esc(r.nm)}<div style="font-size:10.5px;color:#616e7c">${esc(r.cd)}${r.rk ? ' · κωδ. ' + esc(r.rk) : ''} · ${esc(r.how)}</div></td>
            <td>${r.q}</td>
            <td><select class="sel">${optHtml(r)}</select>${r.st !== 'ok' ? '<input type="text" class="sr" placeholder="αναζήτηση στο site (όνομα, κωδικός ή barcode) + Enter">' : ''}</td>
            <td><input type="number" class="qty" min="0" step="1" value="${r.qty}"> <span class="tag">${p ? esc(p.unit) : ''}</span>${p && !p.inStock ? ' <span class="warn">εξαντλ.</span>' : ''}</td></tr>`;
        }).join('')}</tbody></table>`;
    };
    render();
    bd.addEventListener('change', (e) => {
      const tr = e.target.closest('tr[data-i]'); if (!tr) return; const r = R[+tr.dataset.i];
      if (e.target.classList.contains('on')) { r.on = e.target.checked; tr.className = r.on ? r.st : 'off'; }
      if (e.target.classList.contains('sel')) { r.sel = e.target.value; r.mul = 1; r.qty = r.sel ? A.qty(r, bySku[r.sel], 1) : 0; r.user = true; r.st = r.sel ? 'ok' : 'none'; r.how = r.sel ? 'επιλογή σας' : 'δεν βρέθηκε — διαλέξτε'; render(); }
      if (e.target.classList.contains('qty')) { r.qty = Math.max(0, Math.round(Number(e.target.value) || 0)); }
    });
    bd.addEventListener('keydown', async (e) => {
      if (!e.target.classList.contains('sr') || e.key !== 'Enter') return;
      const tr = e.target.closest('tr[data-i]'); const r = R[+tr.dataset.i]; const q = e.target.value.trim(); if (!q) return;
      if (A.search) {   // site χωρίς πλήρη κατάλογο: αναζήτηση στον server
        smsg('Αναζήτηση…');
        (await A.search(q).catch(() => [])).forEach((p) => { if (!bySku[p.sku]) { p._t = toks(p.name); p._v = vol(p.name); prods.push(p); indexP(p); } });
        smsg('');
      }
      const exact = prods.filter((p) => p.sku === q || p.sku.includes(q) && q.length >= 4 || (q.length >= 8 && q in p.bcs)).slice(0, 8).map((p) => [1, p]);
      r.sug = exact.length ? exact : suggest(q, 12); if (r.sug.length) { r.sel = r.sug[0][1].sku; r.mul = 1; r.qty = A.qty(r, bySku[r.sel], 1); r.user = true; r.how = 'επιλογή σας'; }
      render();
    });
    $('spo-s-go').disabled = false;
    $('spo-s-go').onclick = () => addAll();
    // Νέα παραγγελία και όλα βρέθηκαν → μπαίνει κατευθείαν στο καλάθι (χωρίς αποστολή)
    if (fresh && R.every((r) => r.st === 'ok')) setTimeout(() => { smsg('Όλα τα είδη βρέθηκαν — μπαίνουν αυτόματα στο καλάθι…'); addAll(); }, 600);
    else if (fresh) smsg('<b>Ελέγξτε τις κίτρινες/κόκκινες γραμμές</b> και πατήστε «ΠΡΟΣΘΗΚΗ ΣΤΟ ΚΑΛΑΘΙ».');
    async function addAll() {
      const todo = R.filter((r) => r.on && r.sel && r.qty > 0);
      if (!todo.length) { smsg('<span class="warn">Δεν υπάρχει γραμμή για το καλάθι.</span>'); return; }
      $('spo-s-go').disabled = true; smsg('Προσθήκη στο καλάθι…');
      try {
        const qtyBy = {}; todo.forEach((r) => { qtyBy[r.sel] = (qtyBy[r.sel] || 0) + r.qty; });   // ίδιο είδος από 2 γραμμές → άθροισμα
        const res = await A.addToCart(Object.keys(qtyBy).map((sku) => ({ p: bySku[sku], qty: qtyBy[sku] })), $('spo-s-empty').checked);
        R.forEach((r) => { if ((r.user || r.st === 'chk') && r.on && r.sel) learned[r.cd] = r.sel; });
        ls.set('map', learned);
        smsg(`✔ Στο καλάθι: <b>${res.lines}</b> είδη · σύνολο <b>${res.total.toFixed(2)} €</b>${res.net ? ' (χωρίς ΦΠΑ)' : ''}${res.fails.length ? ` · <span class="warn">δεν μπήκαν: ${esc(res.fails.join(', '))}</span>` : ''}`
          + (res.reload ? ' · Η σελίδα ανανεώνεται για να φανεί το καλάθι…' : ' · Ανοίξτε το Καλάθι του site για έλεγχο και αποστολή.'));
        if (res.reload) setTimeout(() => location.reload(), 1800);
      } catch (e) { smsg('<span class="warn">' + esc(e.message || e) + '</span>'); }
      $('spo-s-go').disabled = false;
    }
  }

})();
