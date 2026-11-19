#!/usr/bin/env node
/**
 * Capture a report page with its plotly tooltip open, so the hover content can
 * be reviewed (including by an automated reviewer) without a mouse.
 *
 * Playwright is not installed here and the npm registry is unreachable, so this
 * drives Chrome's DevTools Protocol over Node 22's built-in WebSocket: it
 * locates the top candidate, converts its data coordinates to viewport pixels
 * through plotly's axes, sends a real mouseMoved event, and screenshots the
 * result. Empty `.hoverlayer` text means the tooltip did not open.
 *
 * usage: node capture_report_hover.js "file:///path/report.html#<tab>" out.png <condition-index>
 */
const { spawn } = require("child_process");
const fs = require("fs");
const http = require("http");

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const PORT = 9339;
const [url, out, indexArg] = process.argv.slice(2);
const conditionIndex = parseInt(indexArg || "0", 10);

const get = (path) => new Promise((resolve, reject) => {
  http.get({ host: "127.0.0.1", port: PORT, path }, (res) => {
    let body = "";
    res.on("data", (chunk) => (body += chunk));
    res.on("end", () => resolve(JSON.parse(body)));
  }).on("error", reject);
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const chrome = spawn(CHROME, [
    "--headless=new", "--disable-gpu", "--hide-scrollbars",
    "--remote-debugging-port=" + PORT, "--remote-allow-origins=*",
    "--force-device-scale-factor=2", "--window-size=1500,1700",
    "about:blank",
  ], { stdio: "ignore" });

  let targets = null;
  for (let attempt = 0; attempt < 40 && !targets; attempt++) {
    try { targets = await get("/json/list"); } catch { await sleep(300); }
  }
  if (!targets) { chrome.kill(); throw new Error("chrome devtools endpoint never came up"); }
  const page = targets.find((t) => t.type === "page");
  const ws = new WebSocket(page.webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  ws.addEventListener("message", (event) => {
    const msg = JSON.parse(event.data);
    if (msg.id && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); }
  });
  const send = (method, params = {}) => new Promise((resolve) => {
    const messageId = ++id;
    pending.set(messageId, resolve);
    ws.send(JSON.stringify({ id: messageId, method, params }));
  });
  await new Promise((r) => ws.addEventListener("open", r));

  await send("Page.enable");
  await send("Page.navigate", { url });
  await sleep(14000);   // plotly + the inlined library need a moment

  // Work out the pixel position of the top candidate and send a real mouse move,
  // so the tooltip is produced by the same code path a user triggers.
  const position = await send("Runtime.evaluate", {
    returnByValue: true,
    expression: `(() => {
      const cond = DATA.conditions[${conditionIndex}];
      const gd = document.getElementById("plot");
      if (!gd || !gd._fullLayout) return { error: "plot not ready" };
      const pts = decodePoints(cond).filter((p) => p.bait && !p.crapome
        && p.up >= DATA.thresholds.minUniquePeptides && p.fc >= DATA.thresholds.minFold);
      if (!pts.length) return { error: "no passing prey" };
      const t = pts.sort((a, b) => b.fc - a.fc)[0];
      const box = gd.getBoundingClientRect();
      const ax = gd._fullLayout.xaxis, ay = gd._fullLayout.yaxis;
      const x = box.left + ax._offset + ax.l2p(ax.d2l(t.ctrl));
      const y = box.top + ay._offset + ay.l2p(ay.d2l(t.bait));
      return { x, y, gn: t.gn, name: t.name, fc: t.fc };
    })()`,
  });
  const pos = (position.result && position.result.result && position.result.result.value) || {};
  console.log("  target point:", JSON.stringify(pos));
  if (pos.error) { chrome.kill(); throw new Error(pos.error); }

  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: pos.x, y: pos.y, buttons: 0 });
  await sleep(400);
  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: pos.x + 1, y: pos.y + 1, buttons: 0 });
  await sleep(900);

  await sleep(2500);
  const probe = await send("Runtime.evaluate", {
    returnByValue: true,
    expression: `(() => {
      const layers = document.querySelectorAll(".hoverlayer");
      const texts = [];
      layers.forEach((l) => { const t = l.textContent.trim(); if (t) texts.push(t.slice(0, 160)); });
      return JSON.stringify({ layers: layers.length, texts });
    })()`,
  });
  console.log("  hover layer probe:", probe.result && probe.result.result && probe.result.result.value);
  const shot = await send("Page.captureScreenshot", { format: "png" });
  fs.writeFileSync(out, Buffer.from(shot.result.data, "base64"));
  console.log("  wrote", out);
  ws.close();
  chrome.kill();
  process.exit(0);
})().catch((err) => { console.error("FAILED:", err.message); process.exit(1); });
