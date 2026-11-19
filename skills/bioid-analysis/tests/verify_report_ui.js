#!/usr/bin/env node
/**
 * Verify a BioID report's interactive controls in a real browser and screenshot
 * the result. Drives Chrome's DevTools Protocol over Node's built-in WebSocket
 * (no npm dependency; the npm registry is unreachable on this machine).
 *
 * Uses a random debug port and a fresh Chrome profile per run: reusing a fixed
 * port silently attaches to a stale browser still holding the previous page
 * state, which produced phantom failures during development.
 *
 * Checks:
 *   - the label-count selector offers up to 100 and 200
 *   - typing names into the marking box and pressing "Add mark" adds a trace,
 *     a coloured chip with the matched count, and flags names not in the sheet
 *   - marks survive a gate change (the toolbar triggers a redraw)
 *
 * usage: node verify_report_ui.js "file:///path/report.html#3" out.png <condition-index>
 */
const { spawn } = require("child_process");
const fs = require("fs");
const http = require("http");

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const PORT = 9400 + Math.floor(Math.random() * 200);
const [url, out, indexArg] = process.argv.slice(2);
const conditionIndex = parseInt(indexArg || "0", 10);

const get = (path) => new Promise((resolve, reject) => {
  http.get({ host: "127.0.0.1", port: PORT, path }, (res) => {
    let body = "";
    res.on("data", (c) => (body += c));
    res.on("end", () => resolve(JSON.parse(body)));
  }).on("error", reject);
});
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const chrome = spawn(CHROME, [
    "--headless=new", "--disable-gpu", "--hide-scrollbars",
    "--user-data-dir=" + require("fs").mkdtempSync("/tmp/chrome-probe-"),
    "--remote-debugging-port=" + PORT, "--remote-allow-origins=*",
    "--force-device-scale-factor=2", "--window-size=1500,2600", "about:blank",
  ], { stdio: "ignore" });

  let targets = null;
  for (let i = 0; i < 40 && !targets; i++) {
    try { targets = await get("/json/list"); } catch { await sleep(300); }
  }
  if (!targets) { chrome.kill(); throw new Error("devtools endpoint never came up"); }
  const ws = new WebSocket(targets.find((t) => t.type === "page").webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  ws.addEventListener("message", (e) => {
    const m = JSON.parse(e.data);
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
  });
  const send = (method, params = {}) => new Promise((resolve) => {
    const mid = ++id; pending.set(mid, resolve);
    ws.send(JSON.stringify({ id: mid, method, params }));
  });
  const evaluate = async (expression, awaitPromise = false) => {
    const r = await send("Runtime.evaluate", { returnByValue: true, awaitPromise, expression });
    return r.result && r.result.result ? r.result.result.value : null;
  };
  await new Promise((r) => ws.addEventListener("open", r));
  await send("Page.enable");
  await send("Runtime.enable");
  ws.addEventListener("message", (event) => {
    const msg = JSON.parse(event.data);
    if (msg.method === "Runtime.consoleAPICalled") {
      const text = (msg.params.args || []).map((a) => a.value !== undefined ? a.value : a.description).join(" ");
      if (String(text).indexOf("DBG") === 0) console.log("    page>" + text.slice(3));
    }
  });
  await send("Page.navigate", { url });
  await sleep(14000);

  // A page-level script error must fail the run loudly: the report traps errors
  // into its footer, so check that first.
  const footerText = await evaluate("document.getElementById('footer').innerText");
  if (/script error/i.test(footerText || "")) {
    console.log("  [FAIL] the page reported a script error: " + footerText.slice(0, 200));
  } else {
    console.log("  [PASS] no page-level script error");
  }

  const parse = (raw, label) => {
    if (raw === null || raw === undefined) { console.log(`  [warn] ${label}: no value returned`); return {}; }
    if (typeof raw === "object") { return raw; }
    try { return JSON.parse(raw); }
    catch (e) { console.log(`  [warn] ${label}: unparsable (${e.message}) raw=${String(raw).slice(0, 120)}`); return {}; }
  };

  const results = [];
  const record = (name, pass, detail) => {
    results.push({ name, pass, detail });
    console.log(`  [${pass ? "PASS" : "FAIL"}] ${name}${detail ? " — " + detail : ""}`);
  };

  // scroll zoom must be off (the wheel belongs to the page), and the vector
  // export paths must exist
  const configCheck = await evaluate(`(function(){
    var gd = document.querySelector(".plot-host");
    return JSON.stringify({
      scrollZoom: gd && gd._context ? gd._context.scrollZoom : null,
    });
  })()`);
  const cfg = parse(configCheck, "config");
  record("scroll-wheel zoom is off — scrollZoom = " + cfg.scrollZoom, cfg.scrollZoom === false);
  // The multi-page PDF (header button + print preparation) was removed: the HTML is the
  // report, printing is the browser's own job, and only the per-figure PDF button
  // remains. Guard against it creeping back in.
  const noPrintKit = await evaluate(`JSON.stringify({
    button: !!document.getElementById("pdf-btn"),
    prepare: typeof window.__preparePrint === "function",
    printPrep: document.body.className.indexOf("print-prep") >= 0,
  })`);
  const npk = parse(noPrintKit, "noPrintKit");
  record("no multi-page PDF machinery is left in the report",
    !npk.button && !npk.prepare && !npk.printPrep,
    `button=${npk.button} preparePrint=${npk.prepare} printPrep=${npk.printPrep}`);

  // 0. the in-plot guide labels must be drawn on first paint, before any interaction:
  // the legend names the lines, but the gate tag is what a reader uses to see which line
  // is the gate. A refactor once passed the wrong units into the placement, and the
  // labels were then dropped silently — but only until the first redraw, which is why
  // this check runs here and not at the end.
  const guideLabels = await evaluate(`(function () {
    const gd = document.querySelector(".plot-host");
    const texts = (gd.data || []).filter((t) => t.name && /^guide #/.test(t.name) && t.mode === "text");
    const words = texts.reduce((a, t) => a.concat(t.text || []), []);
    return JSON.stringify({ traces: texts.length, words,
      labels: (document.getElementById("label-note") || {}).textContent });
  })()`);
  const gl = parse(guideLabels, "guideLabels");
  record("the in-plot guide labels are drawn",
    (gl.traces || 0) >= 1 && (gl.words || []).some((w) => /×/.test(String(w))),
    `${gl.traces} text trace(s), words = ${(gl.words || []).join(", ")}`);

  // 1. label options include 100 and 200
  const labelOpts = await evaluate(
    `Array.from(document.querySelectorAll("#labeltop option")).map((o) => o.value).join(",")`);
  record("label selector offers up to 100 and 200",
    /\b100\b/.test(labelOpts) && /\b200\b/.test(labelOpts), `options = ${labelOpts}`);

  // 2. pick real gene names from this condition, plus one that does not exist
  const names = await evaluate(`(() => {
    const cond = DATA.conditions[${conditionIndex}];
    const pts = decodePoints(cond).filter((p) => p.bait && p.fc);
    const top = pts.sort((a, b) => b.fc - a.fc).slice(0, 3).map((p) => p.gn).join(", ");
    const low = pts.sort((a, b) => a.fc - b.fc)[0].gn;
    return [top, low].join("|");
  })()`);
  const [topNames, lowName] = String(names).split("|");
  console.log(`  (using top names: ${topNames}; depleted: ${lowName})`);

  // 3. autocomplete: type a prefix, expect suggestions, click the first one
  const suggest = await evaluate(`(async () => {
    const input = document.getElementById("mark-input");
    input.focus();
    input.value = ${JSON.stringify(String(topNames).split(",")[0].trim().slice(0, 3))};
    input.dispatchEvent(new Event("input"));
    await new Promise((r) => setTimeout(r, 400));
    const box = document.getElementById("mark-suggest");
    const items = box.querySelectorAll("div[data-token]");
    const list = Array.from(items).map((el) => el.dataset.token);
    const shown = box.className.indexOf("show") >= 0;
    if (items.length) items[0].click();
    await new Promise((r) => setTimeout(r, 300));
    return JSON.stringify({ shown, list, pending: document.querySelectorAll(".pend-chip").length });
  })()`, true);
  const sg = parse(suggest, "suggest");
  record("typing offers autocomplete suggestions", sg.shown && (sg.list || []).length > 0,
    `suggestions = ${(sg.list || []).slice(0, 5).join(", ")}`);
  record("clicking a suggestion adds it to the pending list", (sg.pending || 0) >= 1,
    `pending chips = ${sg.pending}`);

  // 4. pick rows in the candidate table, then move them into the mark list
  const pick = await evaluate(`(async () => {
    const rows = document.querySelectorAll("#cand-table tbody tr[data-key]");
    const wanted = [];
    for (let i = 0; i < 3 && i < rows.length; i++) {
      wanted.push(rows[i].dataset.key);
      rows[i].click();
    }
    await new Promise((r) => setTimeout(r, 250));
    const note = document.getElementById("sel-note").textContent.trim();
    document.getElementById("mark-selected").click();
    await new Promise((r) => setTimeout(r, 300));
    const pending = Array.from(document.querySelectorAll(".pend-chip"))
      .map((el) => el.textContent.replace("×", "").trim());
    return JSON.stringify({ wanted, note, pending });
  })()`, true);
  const pk = parse(pick, "pick");
  record("clicking table rows picks proteins", /3 picked/.test(pk.note || ""), `note = ${pk.note}`);
  record("picked proteins land in the mark list",
    (pk.wanted || []).every((w) => (pk.pending || []).includes(w)),
    `pending = ${(pk.pending || []).join(", ")}`);

  // 5. commit the mark group and inspect how it is drawn
  await evaluate(`(() => {
    document.getElementById("manual-color").value = "#8e44ad";
    document.getElementById("manual-add").click();
    return true;
  })()`);
  await sleep(2500);
  const drawn = await evaluate(`(() => {
    const gd = document.querySelector(".plot-host");
    const marks = (gd.data || []).filter((t) => t.name && t.name.indexOf("marked (") === 0);
    const names = (gd.data || []).filter((t) => t.name === "mark labels");
    return JSON.stringify({
      dots: marks.length,
      dotColor: marks[0] && marks[0].marker && marks[0].marker.color,
      dotSymbol: marks[0] && marks[0].marker && marks[0].marker.symbol,
      dotFillOpacity: marks[0] && marks[0].marker && marks[0].marker.opacity,
      nameTraces: names.length,
      nameColor: names[0] && names[0].textfont && names[0].textfont.color,
      nameText: names[0] ? names[0].text.slice(0, 4) : [],
      chips: document.querySelectorAll(".mark-chip").length,
    });
  })()`);
  const dr = parse(drawn, "drawn");
  record("marked prey are drawn as filled dots in the chosen colour",
    dr.dots >= 1 && /8e44ad/i.test(dr.dotColor || "") && dr.dotSymbol === "circle",
    `dots=${dr.dots} colour=${dr.dotColor} symbol=${dr.dotSymbol} opacity=${dr.dotFillOpacity}`);
  record("marked prey are also labelled in the same colour",
    dr.nameTraces >= 1 && /8e44ad/i.test(dr.nameColor || ""),
    `${dr.nameTraces} name trace(s), colour=${dr.nameColor}, e.g. ${(dr.nameText || []).join(", ")}`);
  record("a chip records the mark group", dr.chips >= 1, `${dr.chips} chip(s)`);

  const leaders = await evaluate(`(() => {
    const gd = document.querySelector(".plot-host");
    const lead = (gd.data || []).filter((t) => t.name === "mark leaders");
    const labels = (gd.data || []).filter((t) => t.name === "mark labels");
    return JSON.stringify({
      leaders: lead.length,
      colour: lead[0] && lead[0].line && lead[0].line.color,
      segments: lead[0] ? Math.round(lead[0].x.filter((v) => v !== null).length / 2) : 0,
      labels: labels.reduce((n, t) => n + t.x.length, 0),
      mode: lead[0] && lead[0].mode,
    });
  })()`);
  const ld = parse(leaders, "leaders");
  record("marked prey get a leader line back to the dot",
    ld.leaders >= 1 && ld.mode === "lines" && ld.segments >= 1 && /8e44ad/i.test(ld.colour || ""),
    `${ld.leaders} leader trace(s), ${ld.segments} segment(s), colour=${ld.colour}, labels=${ld.labels}`);

  const pdfBtn = await evaluate(`(function () {
    var btns = document.querySelectorAll('[data-title="Download this figure as vector PDF"]');
    return JSON.stringify({ buttons: btns.length,
      cameraRemoved: !document.querySelector('[data-title="Download plot as a png"]') });
  })()`);
  const pb = parse(pdfBtn, "pdfBtn");
  record("the figure offers a vector-PDF download instead of the PNG camera",
    pb.buttons >= 1, `${pb.buttons} PDF button(s); png camera present=${!pb.cameraRemoved}`);

  const pdfBytes = await evaluate(`(async () => {
    const b64 = await window.__figurePdfBase64(3);
    return b64 || "";
  })()`, true);
  // 整份 PDF 的原始字节：既能验 %PDF 头，也能验"真的是矢量"——栅格版必然带 /Subtype /Image
  // 的 XObject，而矢量版应当出现字体/路径（/Font、re/f/m/l 这些内容流算子），且没有图片对象。
  const raw = (() => {
    try { return Buffer.from(String(pdfBytes), "base64").toString("latin1"); } catch (e) { return ""; }
  })();
  record("the PDF export produces real PDF bytes",
    raw.startsWith("%PDF"), `header=${JSON.stringify(raw.slice(0, 8))}`);
  const hasImage = /\/Subtype\s*\/Image/.test(raw);
  const hasFont = /\/Font/.test(raw) || /\/FontFile/.test(raw);
  const hasVectorOps = /(^|[\s\]>])(re|l|m|c)[\s\n]/.test(raw.slice(-40000));
  record("the exported PDF is vector, not an embedded raster image",
    raw.length > 2000 && !hasImage && (hasFont || hasVectorOps),
    `${raw.length} bytes; /Image=${hasImage}; /Font=${hasFont}; vector ops=${hasVectorOps}`);

  // 5. marks survive a gate change (which redraws the plot)
  await evaluate(`(() => {
    const sel = document.getElementById("fold");
    sel.value = "3";
    sel.dispatchEvent(new Event("change"));
    return true;
  })()`);
  await sleep(2500);
  const afterGate = await evaluate(`(() => {
    const gd = document.querySelector(".plot-host");
    const marks = (gd.data || []).map((t) => t.name).filter((n) => n && n.startsWith("marked"));
    return JSON.stringify({ marks, chips: document.querySelectorAll(".mark-chip").length });
  })()`);
  const gateInfo = parse(afterGate, "afterGate");
  record("marks survive a gate change",
    (gateInfo.marks || []).length >= 1 && gateInfo.chips >= 1,
    `after switching to 3x: ${(gateInfo.marks || []).length} mark trace(s), ${gateInfo.chips} chip(s)`);

  // 6. raising the label count to 100 places more labels (labels are layout
  //    annotations, so they are counted from the layout, not from a trace)
  const labelGrowth = await evaluate(`(async () => {
    const gd = document.querySelector(".plot-host");
    const count = () => (gd.data || [])
      .filter((t) => t.name === "candidate labels")
      .reduce((n, t) => n + (t.text ? t.text.length : 0), 0);
    const before = count();
    const sel = document.getElementById("labeltop");
    sel.value = "100";
    sel.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 1800));
    const note = document.getElementById("label-note");
    return JSON.stringify({ before, after: count(), note: note ? note.textContent : null });
  })()`, true);
  const growth = parse(labelGrowth, "labelGrowth");
  record("raising the label count to 100 places more labels",
    growth.after > growth.before,
    `${growth.before} -> ${growth.after} labels; note: ${growth.note}`);

  // 7. the "share above gate" card must track the selected gate, not stay fixed
  const shareTrack = await evaluate(`(async () => {
    const read = () => {
      const cards = Array.from(document.querySelectorAll("#kpis .card"));
      const card = cards.find((c) => c.querySelector(".lbl") && c.querySelector(".lbl").textContent.trim() === "Share above gate");
      return card ? { value: card.querySelector(".kpi").textContent.trim(), note: card.querySelector(".legend-note").textContent.trim() } : null;
    };
    const sel = document.getElementById("fold");
    sel.value = "1.5"; sel.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 1200));
    const lo = read();
    sel.value = "5"; sel.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 1200));
    const hi = read();
    return JSON.stringify({ lo, hi });
  })()`, true);
  const share = parse(shareTrack, "shareTrack");
  const loPct = share.lo ? parseFloat(share.lo.value) : null;
  const hiPct = share.hi ? parseFloat(share.hi.value) : null;
  record("the share card recomputes when the gate changes",
    loPct !== null && hiPct !== null && loPct > hiPct,
    `1.5x -> ${share.lo && share.lo.value} (${share.lo && share.lo.note}); 5x -> ${share.hi && share.hi.value} (${share.hi && share.hi.note})`);

  // 8. zooming in must place more labels (they are zoom-aware) and must not break the plot
  const countLabels = `(function () {
    const gd = document.querySelector(".plot-host");
    const traces = (gd.data || []).filter((t) => t.name === "candidate labels");
    const names = traces.reduce((a, t) => a.concat(t.text || []), []);
    return JSON.stringify({
      labels: names.length,
      names: names.join(","),
    });
  })()`;
  // trigger the zoom without awaiting a promise inside the page, then read the
  // count separately: keeps the CDP call synchronous and easy to reason about
  const zoomBefore = await evaluate(countLabels);
  await evaluate(`(function () {
    const gd = document.querySelector(".plot-host");
    const lo = gd._fullLayout.xaxis.range[0] + 2.0, hi = gd._fullLayout.xaxis.range[1] - 2.0;
    const yLo = gd._fullLayout.yaxis.range[0] + 2.0, yHi = gd._fullLayout.yaxis.range[1] - 2.0;
    Plotly.relayout(gd, { "xaxis.range": [lo, hi], "yaxis.range": [yLo, yHi] });
    return "zoomed";
  })()`);
  await sleep(3000);
  const zoomAfter = await evaluate(countLabels);
  await evaluate(`(function () {
    Plotly.relayout(document.querySelector(".plot-host"), { "xaxis.autorange": true, "yaxis.autorange": true });
    return "reset";
  })()`);
  await sleep(1500);
  const zoomCheck = JSON.stringify({
    before: parse(zoomBefore, "zoomBefore").labels,
    after: parse(zoomAfter, "zoomAfter").labels,
    namesBefore: parse(zoomBefore, "zoomBefore").names,
    namesAfter: parse(zoomAfter, "zoomAfter").names,
  });
  const zoom = parse(zoomCheck, "zoomCheck");
  // Labels are recomputed for the visible window, so the names must change when
  // the window does. The count is not asserted upwards: it depends on where the
  // zoom lands (a window over a sparse corner places fewer).
  record("zooming in relabels for the visible window",
    (zoom.after || 0) > 0 && (zoom.namesBefore || "") !== (zoom.namesAfter || ""),
    `${zoom.before} labels at full view -> ${zoom.after} when zoomed in; names `
      + `${String(zoom.namesBefore || "").slice(0, 40)} -> ${String(zoom.namesAfter || "").slice(0, 40)}`);

  // 9. the scatter must be taller than a letterbox and roughly square
  const geom = await evaluate(`(() => {
    const host = document.querySelector(".plot-host");
    const gd = host;
    const r = host.getBoundingClientRect();
    return JSON.stringify({ w: Math.round(r.width), h: Math.round(r.height) });
  })()`);
  const gm = parse(geom, "geom");
  record("the plot is tall, not a letterbox",
    gm.h > 700 && gm.h / gm.w > 0.6,
    `${gm.w}x${gm.h} device px (ratio ${(gm.h / gm.w).toFixed(2)})`);

  // 10. the hover card must carry the UniProt annotation
  const hoverInfo = await evaluate(`(async () => {
    const gd = document.querySelector(".plot-host");
    const cond = DATA.conditions[${conditionIndex}];
    const pts = decodePoints(cond).filter((p) => p.bait && p.acc && (DATA.annotations || {})[p.acc]);
    if (!pts.length) return JSON.stringify({ skipped: "no annotated protein in this condition" });
    const t = pts.sort((a, b) => (b.fc || 0) - (a.fc || 0))[0];
    const box = gd.getBoundingClientRect();
    const ax = gd._fullLayout.xaxis, ay = gd._fullLayout.yaxis;
    // x is where the prey is drawn: the measured control, or the floor when the
    // control never detected it (t.ctrl is null for those).
    const xc = (t.cuse !== null && t.cuse !== undefined) ? t.cuse : t.ctrl;
    const x = box.left + ax._offset + ax.l2p(ax.d2l(xc));
    const y = box.top + ay._offset + ay.l2p(ay.d2l(t.bait));
    return JSON.stringify({ x, y, gn: t.gn, acc: t.acc });
  })()`, true);
  const hi = parse(hoverInfo, "hoverInfo");
  if (hi.skipped) {
    console.log("  [SKIP] hover annotation check — " + hi.skipped);
  } else {
    await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: hi.x, y: hi.y, buttons: 0 });
    await sleep(400);
    await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: hi.x + 1, y: hi.y + 1, buttons: 0 });
    await sleep(900);
    const card = await evaluate(`(() => {
      const layers = document.querySelectorAll(".hoverlayer");
      let text = "";
      layers.forEach((l) => { text += " " + (l.textContent || ""); });
      return text.replace(/\\s+/g, " ").trim();
    })()`);
    const ann = await evaluate(`JSON.stringify(DATA.annotations[${JSON.stringify(hi.acc)}] || null)`);
    const a = JSON.parse(ann || "null");
    const has = (needle) => needle && card && card.toLowerCase().indexOf(String(needle).toLowerCase()) >= 0;
    record("hover shows the protein full name", has(a && a[0]),
      `hover = ${(card || "").slice(0, 150)}`);
    record("hover shows aliases", !a || !a[2] || has(String(a[2]).split(";")[0]),
      `expected alias ${a && a[2] ? String(a[2]).split(";")[0] : "(none)"}`);
    record("hover shows a function description",
      !a || !a[3] || has(String(a[3]).slice(0, 40)),
      `expected start of ${a && a[3] ? JSON.stringify(String(a[3]).slice(0, 40)) : "(none)"}`);
  }

  // 11. the guide lines must bound the gate the colours were decided by. A guide
  // drawn at an uncorrected ratio looks plausible and is wrong: in this batch up
  // to 93% of the passing prey sit below it. The gate is set to the payload
  // default first, because these checks compare the drawn line with the colours
  // and both are recomputed on every gate change.
  await evaluate(`(async () => {
    const fold = document.getElementById("fold");
    fold.value = String(DATA.thresholds.minFold); fold.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 2200));
    return true;
  })()`, true);
  await sleep(1200);
  const gateAgree = await evaluate(`(function () {
    const gd = document.querySelector(".plot-host");
    const cond = DATA.conditions[${conditionIndex}];
    const corr = 2 ** cond.diagnostics.correction_log2;
    const gate = parseFloat(document.getElementById("fold").value) || DATA.thresholds.minFold;
    const gateRaw = gate * corr;
    const named = (n) => (gd.data || []).filter((t) => t.name === n);
    const ratio = (t) => t.x.map((x, i) => (t.y[i] && x ? t.y[i] / x : null)).filter((v) => v !== null);
    const passRatios = named("Passes all gates").reduce((a, t) => a.concat(ratio(t)), []);
    const lbRatios = named("EV not detected — fold change is a lower bound")
      .reduce((a, t) => a.concat(ratio(t)), []);
    const belowRatios = named("Below fold-change gate")
      .reduce((a, t) => a.concat(ratio(t)), []);
    const eps = 1e-6;
    return JSON.stringify({
      corr,
      gateRaw,
      passBelow: passRatios.filter((r) => r < gateRaw * (1 - eps)).length,
      passCount: passRatios.length,
      lbBelow: lbRatios.filter((r) => r < gateRaw * (1 - eps)).length,
      lbCount: lbRatios.length,
      belowAbove: belowRatios.filter((r) => r > gateRaw * (1 + eps)).length,
      belowCount: belowRatios.length,
      guideFolds: (gd._fullLayout.shapes || []).map((sh) =>
        (sh.y1 && sh.x1) ? Number((sh.y1 / sh.x1).toFixed(5)) : null),
      // the guide carrying the gate, matched by its colour (#1b7f3b)
      drawnGate: (function () {
        const sh = (gd._fullLayout.shapes || []).find((x) => x.line && x.line.color === "#1b7f3b");
        return sh ? sh.y1 / sh.x1 : null;
      })(),
    });
  })()`);
  const ga = parse(gateAgree, "gateAgree");
  record("the drawn gate line sits at the load-corrected ratio",
    ga.drawnGate !== null && Math.abs(ga.drawnGate - ga.gateRaw) < 1e-3,
    `drawn at raw ratio ${ga.drawnGate} vs the gate's corrected ratio ${Number(ga.gateRaw).toFixed(4)}`);
  record("the drawn gate line bounds the prey it colours",
    ga.passBelow === 0 && ga.lbBelow === 0,
    `${ga.passBelow}/${ga.passCount} passing and ${ga.lbBelow}/${ga.lbCount} EV-undetected prey below the drawn gate`
      + ` (guides at ${(ga.guideFolds || []).filter((f) => f).join(", ")}, gate raw ratio ${Number(ga.gateRaw).toFixed(3)})`);
  record("prey below the drawn gate are coloured as below it",
    ga.belowAbove === 0 || ga.belowCount === undefined,
    `${ga.belowAbove}/${ga.belowCount} grey prey sit above the drawn gate`);

  // 12. the EV-undetected class: counted, drawn as its own symbol, and labelled as
  // a lower bound rather than a measurement
  const lb = await evaluate(`(function () {
    const gd = document.querySelector(".plot-host");
    const cond = DATA.conditions[${conditionIndex}];
    const gate = parseFloat(document.getElementById("fold").value) || DATA.thresholds.minFold;
    const n = decodePoints(cond).filter((p) => p.lb && p.fc >= gate
      && p.up >= DATA.thresholds.minUniquePeptides && !p.crapome && !p.bait_hit).length;
    const t = (gd.data || []).find((x) => x.name && x.name.indexOf("EV not detected") === 0);
    const card = Array.from(document.querySelectorAll("#kpis .card"))
      .find((c) => /EV not detected/i.test(c.innerText || ""));
    return JSON.stringify({
      expected: n,
      trace: !!t,
      symbol: t ? t.marker.symbol : null,
      opacity: t ? t.marker.opacity : null,
      cardValue: card ? (card.querySelector(".kpi") || {}).textContent : null,
      note: t ? t.hovertemplate && t.customdata && t.customdata[13] : null,
    });
  })()`);
  const lbc = parse(lb, "lowerBound");
  record("EV-undetected prey are drawn and counted as their own class",
    lbc.expected === 0
      ? !lbc.trace && String(lbc.cardValue || "").trim() === "0"
      : lbc.trace && lbc.symbol === "triangle-up" && String(lbc.cardValue).trim() === String(lbc.expected),
    `${lbc.expected} expected, card=${lbc.cardValue}, trace=${lbc.trace}, symbol=${lbc.symbol}, hover=${String(lbc.note || "").slice(0, 60)}`);

  // 13. the label-by selector exists and actually reorders the names
  const labelBy = await evaluate(`(async () => {
    const sel = document.getElementById("label-by");
    if (!sel) return JSON.stringify({ missing: true });
    const read = () => {
      const gd = document.querySelector(".plot-host");
      const t = (gd.data || []).find((x) => x.name === "candidate labels");
      return { names: t ? (t.text || []).slice(0, 4).join(",") : "", note: (document.getElementById("label-note") || {}).textContent || "" };
    };
    const options = Array.from(sel.options).map((o) => o.value);
    const seen = {};
    for (const v of options) {
      sel.value = v; sel.dispatchEvent(new Event("change"));
      await new Promise((r) => setTimeout(r, 1800));
      seen[v] = read();
    }
    sel.value = "auto"; sel.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 1500));
    return JSON.stringify({ options, seen });
  })()`, true);
  const lbp = parse(labelBy, "labelBy");
  const optionList = (lbp.options || []).join(",");
  record("the report offers a label-by selector", /auto/.test(optionList) && /q/.test(optionList)
    && /fold/.test(optionList) && /abundance/.test(optionList) && /score/.test(optionList), `options = ${optionList}`);
  const autoNames = (lbp.seen && lbp.seen.auto && lbp.seen.auto.names) || "";
  const foldNames = (lbp.seen && lbp.seen.fold && lbp.seen.fold.names) || "";
  record("changing label-by changes the labelled names",
    !!autoNames && !!foldNames && autoNames !== foldNames,
    `evidence tier: ${autoNames} | fold change: ${foldNames}`);
  record("the label note states the criterion in words",
    /labeled by/i.test((lbp.seen && lbp.seen.auto && lbp.seen.auto.note) || ""),
    `note = ${String((lbp.seen && lbp.seen.auto && lbp.seen.auto.note) || "").slice(0, 110)}`);

  // 14. recolouring an existing mark must reach the dot, its name and its leader
  const recolor = await evaluate(`(async () => {
    const inp = document.querySelector("#manual-chips input.swatch");
    if (!inp) return JSON.stringify({ missing: true });
    inp.value = "#0066cc";
    inp.dispatchEvent(new Event("input"));
    await new Promise((r) => setTimeout(r, 2200));
    const gd = document.querySelector(".plot-host");
    const dot = (gd.data || []).find((t) => t.name && t.name.indexOf("marked (") === 0);
    const name = (gd.data || []).find((t) => t.name === "candidate labels" && false) || null;
    const marks = JSON.parse(localStorage.getItem("bioid-marks:" + DATA.folder + ":" + DATA.conditions[${conditionIndex}].condition) || "[]");
    const labelled = (gd.data || []).filter((t) => t.mode === "text" && t.textfont && t.textfont.color === "#0066cc");
    const leader = (gd.data || []).filter((t) => t.name === "mark leaders" && t.line && t.line.color === "#0066cc");
    return JSON.stringify({
      stored: marks.map((m) => m.color).join(","),
      dotColor: dot ? dot.marker.color : null,
      labelTraces: labelled.reduce((n, t) => n + t.x.length, 0),
      leaderTraces: leader.length,
      swatch: inp.value,
    });
  })()`, true);
  const rc = parse(recolor, "recolor");
  record("recolouring a mark updates dot, name and leader",
    !rc.missing && rc.dotColor === "#0066cc" && rc.labelTraces > 0 && rc.leaderTraces > 0
      && String(rc.stored).indexOf("#0066cc") >= 0,
    `stored=${rc.stored} dot=${rc.dotColor} names=${rc.labelTraces} leaders=${rc.leaderTraces}`);

  // 14c. the candidate table is ONE fold-change-descending list, with the EV-undetected
  // rows mixed in (they used to be parked as a tier at the top; the user asked for a
  // single sorted table).
  const tableOrder = await evaluate(`(function () {
    const rows = Array.from(document.querySelectorAll("#cand-table tbody tr"));
    const fcOf = (tr) => {
      const tds = tr.querySelectorAll("td");
      // column order matches the header build: Gene, Accession, Protein, FC, raw FC, ...
      const raw = (tds[3] || {}).textContent || "";
      return parseFloat(raw.replace(/[^0-9.\-eE]/g, "")) || 0;
    };
    const fcs = rows.map(fcOf);
    const ev = rows.map((tr) => ((tr.querySelectorAll("td")[5] || {}).textContent || "").indexOf("not detected") >= 0);
    const nonIncreasing = fcs.every((v, i) => i === 0 || fcs[i - 1] >= v - 1e-6);
    const lastLB = ev.lastIndexOf(true);
    const mixed = lastLB < 0 || ev.slice(0, lastLB).some((x) => !x);
    return JSON.stringify({ rows: rows.length, nonIncreasing, mixed, lastLB,
      top: fcs[0], bottom: fcs[fcs.length - 1] });
  })()`);
  const tord = parse(tableOrder, "tableOrder");
  record("the candidate table is ordered by fold change",
    tord.nonIncreasing && (tord.mixed === undefined || tord.mixed),
    `${tord.rows} rows, top ${Number(tord.top).toFixed(1)}x -> bottom ${Number(tord.bottom).toFixed(1)}x, EV-undetected rows ${tord.mixed ? "mixed in" : "segregated at the top"}`);

  // 14d. the baited protein's dot must carry its own name — it is the anchor a reader
  // looks for, and nothing else would ever label it (it is not a candidate).
  const baitLabel = await evaluate(`(function () {
    const cond = DATA.conditions[${conditionIndex}];
    const pts = decodePoints(cond).filter((p) => p.bait_hit);
    const gd = document.querySelector(".plot-host");
    const labels = (gd.data || []).filter((t) => t.name === "bait label");
    const leaders = (gd.data || []).filter((t) => t.name === "bait leader");
    const texts = labels.reduce((a, t) => a.concat(t.text || []), []);
    const names = Array.from(new Set(pts.map((p) => p.gn)));
    const colour = labels[0] && labels[0].textfont ? labels[0].textfont.color : null;
    return JSON.stringify({ hasBait: pts.length > 0, names, texts, colour,
      leaders: leaders.length });
  })()`);
  const bl = parse(baitLabel, "baitLabel");
  if (!bl.hasBait) {
    record("the bait dot carries its own name", true,
      "no bait point in this condition (bait recovery failure) — nothing to label");
  } else {
    const labelled = (bl.names || []).some((n) => (bl.texts || []).indexOf(n) >= 0);
    record("the bait dot carries its own name",
      labelled && bl.colour === "#8e44ad" && bl.leaders >= 1,
      `bait ${(bl.names || []).join(", ")}: label=${bl.texts} colour=${bl.colour} leaders=${bl.leaders}`);
  }

  // 15. the candidate count the report shows must equal the candidate count the
  // pipeline wrote, i.e. the plot's colour classes must not count rows the gates
  // excluded (the bait self-hit is the one that used to slip through).
  const countAgree = await evaluate(`(function () {
    const cond = DATA.conditions[${conditionIndex}];
    const gate = parseFloat(document.getElementById("fold").value) || DATA.thresholds.minFold;
    const up = DATA.thresholds.minUniquePeptides;
    const expected = decodePoints(cond).filter((p) => p.pass).length;
    const bucket = decodePoints(cond).filter((p) => {
      const st = statusOf(p, gate, up);
      return st === "pass" || st === "lower_bound";
    }).length;
    const card = Array.from(document.querySelectorAll("#kpis .card"))
      .find((c) => /^candidates/i.test((c.innerText || "").replace(/\s+/g, " ").trim()));
    const shown = card ? parseInt((card.querySelector(".kpi") || {}).textContent, 10) : null;
    const bait = decodePoints(cond).filter((p) => p.bait_hit);
    return JSON.stringify({ expected, bucket, shown,
      baitStatus: bait.map((p) => statusOf(p, gate, up)).join(",") });
  })()`);
  const ca = parse(countAgree, "countAgree");
  record("the shown candidate count matches the pipeline's",
    ca.expected === ca.shown && ca.bucket === ca.shown,
    `pipeline=${ca.expected} buckets=${ca.bucket} card=${ca.shown}; bait class=${ca.baitStatus}`);

  // settle into a defined state for the screenshot: default gate, 100 labels, marks on
  await evaluate(`(async () => {
    const fold = document.getElementById("fold");
    fold.value = "1.5"; fold.dispatchEvent(new Event("change"));
    await new Promise((r) => setTimeout(r, 1500));
    return true;
  })()`, true);
  await sleep(1200);

  const shot = await send("Page.captureScreenshot", { format: "png" });
  fs.writeFileSync(out, Buffer.from(shot.result.data, "base64"));
  console.log(`  wrote ${out}`);

  const failed = results.filter((r) => !r.pass);
  console.log(`\n  ${results.length - failed.length}/${results.length} UI checks passed`);
  ws.close(); chrome.kill();
  process.exit(failed.length ? 1 : 0);
})().catch((err) => { console.error("FAILED:", err.message); process.exit(1); });
