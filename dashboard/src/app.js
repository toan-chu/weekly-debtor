/* Credit Dashboard — đọc file Excel tuần do tool weekly-debtor sinh ra, vẽ trong trình duyệt.
   Không tính lại tiền: mọi con số lấy từ file. Chỉ tính lại NHÓM RỦI RO khi kế toán đã sửa
   lý do trễ hạn trong sheet Receivables (cùng luật với sheet AR Risk / Methodology).
   Kiểu Looker Studio: mỗi thẻ chọn được kiểu biểu đồ hoặc bảng; bấm vào một phần biểu đồ
   để lọc (Sales, nhóm rủi ro, tuổi nợ).
   WR-07 (user 09/10): mỗi tab là MỘT TRANG riêng, không cuộn dài. Bấm một khách / một Sales
   thì vào trang chi tiết của họ (drill-through); bộ lọc giữ nguyên khi chuyển tab nên các tab khác
   chỉ còn phần liên quan. Địa chỉ trang nằm trên thanh địa chỉ (#sales?s=...) nên nút Back chạy. */
(function () {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const fmt = new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 0 });
  const n0 = (v) => (typeof v === "number" ? v : 0);
  const money = (v) => (typeof v === "number" && Math.abs(v) >= 0.5 ? fmt.format(Math.round(v)) : "—");
  const short = (v) => {
    const a = Math.abs(v);
    if (a >= 1e9) return (v / 1e9).toLocaleString("vi-VN", { maximumFractionDigits: 2 }) + " tỷ";
    if (a >= 1e6) return (v / 1e6).toLocaleString("vi-VN", { maximumFractionDigits: 0 }) + " tr";
    return fmt.format(v);
  };
  const pct = (a, b) => (b ? ((a / b) * 100).toFixed(1).replace(".", ",") + "%" : "—");
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const BUCKETS = ["Current", "1 - 30", "31 - 60", "61 - 90", "91 - 120", "120+"];
  const BUCKET_VI = ["Chưa đến hạn", "Quá hạn 1–30 ngày", "31–60 ngày", "61–90 ngày", "91–120 ngày", "Trên 120 ngày"];
  // LUẬT MÀU (user 08/10): đỏ và các màu lân cận (cam, vàng cam) = TIÊU CỰC;
  // xanh lá = TÍCH CỰC; xanh dương, tím = TRUNG TÍNH (chỉ để phân biệt, không khen chê).
  const POS = "#1baf7a", NEU = "#2a78d6", NEU2 = "#7a3fc2", NEG = "#e34948";
  // tuổi nợ: chưa đến hạn là tốt (xanh lá), càng quá hạn càng đỏ đậm
  const AGE_COLORS = ["#1baf7a", "#f2c14e", "#f08a3c", "#e35d4f", "#c53a3a", "#8e1b2b"];
  // màu phân biệt (Sales, khách, vendor): chỉ họ xanh dương / xanh lá / tím — không dùng đỏ, cam.
  // Đã chạy bộ kiểm tra màu dataviz: đạt cả mù màu; quá 7 mục gộp "Khác" (xám).
  const CAT = ["#2a78d6", "#1baf7a", "#4a3aa7", "#0e9aa7", "#7a3fc2", "#5b8def", "#008300"];
  const OTHER = "#a39fae";
  const MUC_FIN = ["Bình thường", "Cảnh báo", "Rủi ro cao"];   // WR-13, cùng wr/thuonghieu.MUC_FIN
  const TIER_COLOR = { N4: "#8e1b2b", N3: "#e34948", N2: "#f08a3c", N1: "#f2c14e", X: "#2a78d6", "0": "#1baf7a" };
  const AUTO = "[TỰ ĐỘNG]";
  const KW_X = ["cấn trừ", "bù trừ", "thanh toán hộ", "supplier của mình"];
  const KW_N4 = ["đang kiện", "mất liên lạc", "khó đòi", "công nợ xấu", "phá sản"];

  let D = null;
  const F = { sales: "", tier: "", bucket: -1, q: "", cust: "" };   // bộ lọc chung, đi theo qua các tab
  const PAGES = ["tong-quan", "khach", "sales", "rui-ro", "tra", "tien"];
  let PAGE = "tong-quan";
  const VIEW = {};                                         // kiểu biểu đồ đang chọn của từng thẻ

  // ---------------- nạp file ----------------
  const drop = $("#drop"), input = $("#file");
  drop.addEventListener("click", () => input.click());
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") input.click(); });
  $("#btnFile").addEventListener("click", () => input.click());
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) load(f); });
  input.addEventListener("change", () => { if (input.files[0]) load(input.files[0]); input.value = ""; });
  $("#btnPrint").addEventListener("click", () => window.print());
  let hen = 0;   // biểu đồ vẽ theo bề rộng thật của thẻ: đổi cỡ cửa sổ thì vẽ lại
  window.addEventListener("resize", () => { clearTimeout(hen); hen = setTimeout(() => D && render(), 150); });
  window.addEventListener("beforeprint", () => { if (D) { document.body.classList.add("printing"); render(); } });
  window.addEventListener("afterprint", () => { document.body.classList.remove("printing"); if (D) render(); });
  window.addEventListener("hashchange", () => { if (!D) return; readHash(); syncControls(); render(); window.scrollTo(0, 0); });

  function load(file) {
    const r = new FileReader();
    r.onload = () => {
      try {
        const wb = XLSX.read(new Uint8Array(r.result), { type: "array" });
        D = parse(wb);
        D.file = file.name;
        $("#err").hidden = true;
        $("#app").hidden = false;
        drop.classList.add("mini");
        drop.hidden = true;                 // đã có nút "Nạp file khác" trên thanh đầu trang
        $("#btnPrint").disabled = false;
        setupFilters();
        readHash();
        syncControls();
        render();
        navReady();
      } catch (e) {
        $("#err").textContent = "Không đọc được file: " + e.message + ". Hãy chọn file …_Credit_Report_….xlsx do tool sinh ra.";
        $("#err").hidden = false;
        console.error(e);
      }
    };
    r.readAsArrayBuffer(file);
  }

  // ---------------- đọc Excel ----------------
  const rows = (wb, name) => {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error("thiếu sheet " + name);
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  };
  const findRow = (R, pred) => R.findIndex((r) => r && pred(r.map((v) => String(v ?? "").trim())));

  function parse(wb) {
    const d = {};
    // Summary: cột F là khoá máy đọc
    const S = rows(wb, "Summary");
    d.period = String(S[2]?.[0] || "");
    { const m = d.period.match(/At\s+(\d{1,2})\s+(\w{3})\s+(\d{4})/), M = "JanFebMarAprMayJunJulAugSepOctNovDec";
      ASOF = m ? new Date(+m[3], M.indexOf(m[2]) / 3, +m[1]) : new Date(); }
    d.now = {}; d.last = {};
    S.forEach((r) => { if (r && typeof r[5] === "string" && /^[a-z_0-9]+$/.test(r[5])) { d.now[r[5]] = r[2]; d.last[r[5]] = r[1]; } });
    const hi = findRow(S, (r) => r[0] === "Week" && r[1] === "Period end");
    d.history = [];
    if (hi >= 0) {
      const h = S[hi].map((v) => String(v ?? ""));
      for (let i = hi + 1; i < S.length && S[i] && S[i][0]; i++) {
        const o = {}; h.forEach((k, j) => (o[k] = S[i][j])); d.history.push(o);
      }
    }
    // Receivables: dòng khách + dòng con (job) bắt đầu bằng └
    const R = rows(wb, "Receivables");
    const ri = findRow(R, (r) => r.includes("Code") && r.includes("Total"));
    // WR-11 PA1: cột "Reason last week (read only)" — tìm theo tiêu đề (file cũ không có thì bỏ qua)
    const H0 = (R[ri] || []).map((v) => String(v ?? "").trim());
    const iLw = H0.findIndex((v) => v.startsWith("Reason last week"));
    // WR-12: từ 29/10 lý do tuần trước ở N, tuần này ở O, Đối chiếu ở P — tìm theo tiêu đề (file cũ: 13 / 14)
    const iLd = H0.indexOf("Reason for late payment") >= 0 ? H0.indexOf("Reason for late payment") : 13;
    const iDc = H0.findIndex((v) => v.startsWith("Đối chiếu")) >= 0 ? H0.findIndex((v) => v.startsWith("Đối chiếu")) : 14;
    // WR-13: thêm cột "Risk level" + 3 ô Sales → đọc mọi cột theo tiêu đề (file cũ: vị trí cũ)
    const colR = (name, def) => { const k = H0.indexOf(name); return k >= 0 ? k : def; };
    const iTerm = colR("Credit Term", 3), iTot = colR("Total", 4), iB0 = colR("Current", 5);
    const iSal = colR("Salesman", 11), iJob = colR("Job Number", 12);
    const iDg = ["Adverse info? (Sales)", "Customer responds? (Sales)", "Broken promises / often late? (Sales)"].map((n) => H0.indexOf(n));
    d.cust = [];
    let cur = null;
    for (let i = ri + 1; i < R.length; i++) {
      const r = R[i]; if (!r) continue;
      const name = String(r[2] ?? "").trim();
      if (String(r[1] ?? "").trim().toUpperCase() === "TOTAL") break;
      if (name.startsWith("└")) {
        if (cur) cur.kids.push({ label: name.replace(/^└\s*/, ""), amount: n0(r[iTot]), sales: r[iSal] || "", job: r[iJob] || "",
                                 b: BUCKETS.map((_, k) => n0(r[iB0 + k])), status: String(r[iDc] ?? "") });
        continue;
      }
      if (!r[1] || typeof r[iTot] !== "number") continue;
      cur = { code: String(r[1]).trim(), name, term: r[iTerm] || "", total: r[iTot], b: BUCKETS.map((_, k) => n0(r[iB0 + k])),
              sales: r[iSal] || "", job: r[iJob] || "", reason: String(r[iLd] ?? ""), dc: String(r[iDc] ?? ""),
              dg: iDg.map((k, j) => (k >= 0 && String(r[k] ?? "").trim()) || ["Không", "Có", "Không"][j]),
              reasonLw: iLw >= 0 ? String(r[iLw] ?? "") : "", kids: [], inv: [] };
      cur.overdue = cur.b.slice(1).reduce((a, b) => a + b, 0);
      cur.oldest = cur.b.reduce((m, v, k) => (v > 0.5 ? k : m), 0);
      d.cust.push(cur);
    }
    const byCode = Object.fromEntries(d.cust.map((c) => [c.code.toUpperCase(), c]));
    // AR Risk
    const K = rows(wb, "AR Risk");
    const ki = findRow(K, (r) => r.includes("Risk Tier"));
    const kh = K[ki].map((v) => String(v ?? "").trim());
    const col = (n) => kh.indexOf(n);
    for (let i = ki + 1; i < K.length; i++) {
      const r = K[i]; if (!r || !r[1] || typeof r[0] !== "number") continue;
      const c = byCode[String(r[1]).trim().toUpperCase()]; if (!c) continue;
      // WR-08: S1–S6, nhóm, xu hướng, việc cần làm trong file là CÔNG THỨC (có thể chưa có giá trị nếu file
      // chưa mở bằng Excel) → dashboard tự tính lại từ các cột số, cùng luật Head of FIN
      c.lwTotal = n0(r[col("Last week total")]); c.lwOver = n0(r[col("Last week overdue")]);
      c.lw61 = n0(r[col("Last week 61+")]); c.bySales = String(r[col("Jobs by Salesman")] ?? "");
      const pd = r[col("Promise date")];
      c.pdFile = typeof pd === "number" ? new Date(Math.round((pd - 25569) * 864e5)) : pd instanceof Date ? pd : null;
      c.maxQh = col("Max days overdue") >= 0 ? n0(r[col("Max days overdue")]) : null;
      c.opening = n0(r[col("Opening (Misa)")]); c.inc = n0(r[col("Increase")]); c.dec = n0(r[col("Decrease")]);
    }
    d.cust.forEach(tierOf);
    // WR-08: sheet Contracts — cảnh báo hợp đồng / credit term (đọc từ Client Master)
    d.contracts = [];
    if (wb.Sheets["Contracts"]) {
      const H = rows(wb, "Contracts"), hi = findRow(H, (r) => r.includes("Alert"));
      if (hi >= 0) for (let i = hi + 1; i < H.length; i++) {
        const r = H[i]; if (!r || typeof r[0] !== "number") continue;
        const han = typeof r[8] === "number" ? ngayVN(new Date(Math.round((r[8] - 25569) * 864e5))) : String(r[8] ?? "");
        const a = { loai: String(r[1] || ""), code: String(r[2] ?? "").trim(), name: String(r[3] || ""), sales: String(r[4] || ""),
                    balance: n0(r[5]), hd: String(r[6] || ""), term: String(r[7] || ""), han, left: r[9], note: String(r[10] || "") };
        a.muc = a.loai.startsWith("Hợp đồng đã hết hạn") ? "do" : /≤30|Chưa có credit|Chưa có trong (Client Master|file credit term)|vừa đổi/.test(a.loai) ? "vang" : "";
        d.contracts.push(a);
        const c = byCode[a.code.toUpperCase()]; if (c) (c.alerts = c.alerts || []).push(a);
      }
    }
    // Invoices
    if (wb.Sheets["Invoices"]) {
      const V = rows(wb, "Invoices");
      const vi = findRow(V, (r) => r.includes("Invoice No."));
      for (let i = vi + 1; i < V.length; i++) {
        const r = V[i]; if (!r || typeof r[0] !== "number") continue;
        const c = byCode[String(r[1] ?? "").trim().toUpperCase()]; if (!c) continue;
        c.inv.push({ no: r[3], date: r[4], due: r[6], days: r[7], bucket: r[8], amount: n0(r[9]), note: r[11] || "" });
      }
    }
    // Payable
    d.vendors = [];
    const P = rows(wb, "Payable");
    const pi = findRow(P, (r) => r.includes("Vendor's Name"));
    if (pi >= 0) for (let i = pi + 1; i < P.length; i++) {
      const r = P[i]; if (!r || typeof r[0] !== "number") continue;
      // FIN 09/10: sheet Payable theo mẫu FIN — No · Code · Vendor's Name · Current · Total
      d.vendors.push({ code: r[1], name: r[2], close: n0(r[4]), paid: 0 });
    }
    // Cash Flow — phần A (khách trả) và phần B (đã trả vendor trong tuần)
    d.payers = [];
    const C = rows(wb, "Cash Flow");
    const bi = findRow(C, (r) => r.includes("Vendor's Name") && r.includes("Paid this week"));
    if (bi >= 0) {
      const byV = Object.fromEntries(d.vendors.map((v) => [String(v.code), v]));
      for (let i = bi + 1; i < C.length; i++) {
        const r = C[i]; if (!r || typeof r[0] !== "number") break;
        const v = byV[String(r[1])];
        if (v) v.paid = n0(r[3]); else d.vendors.push({ code: r[1], name: r[2], close: n0(r[4]), paid: n0(r[3]) });
      }
    }
    const ci = findRow(C, (r) => r.includes("Customer's Name") && r.includes("Bank"));
    if (ci >= 0) for (let i = ci + 1; i < C.length; i++) {
      const r = C[i]; if (!r || typeof r[0] !== "number") break;
      d.payers.push({ code: r[1], name: r[2], bank: n0(r[3]), cash: n0(r[4]), offset: n0(r[5]), fx: n0(r[6]), other: n0(r[7]), total: n0(r[8]) });
    }
    return d;
  }

  // S7 lỡ hẹn (user 13/10): ngày hẹn trả muộn nhất ghi trong lý do — cùng luật với tindung.ngay_hen
  let ASOF = new Date();
  function ngayHen(text) {
    let t = String(text || "").toLowerCase(); const ra = [];
    if (!t.trim() || t.startsWith(AUTO.toLowerCase())) return null;
    const them = (d, m, y, cuoi) => {
      d = +d; m = +m; if (m < 1 || m > 12) return;
      let nam = y ? +y + (+y < 100 ? 2000 : 0) : ASOF.getFullYear();
      if (!y && new Date(nam, m - 1, 1) - ASOF > 183 * 864e5) nam--;
      else if (!y && ASOF - new Date(nam, m - 1, 28) > 183 * 864e5) nam++;
      const x = new Date(nam, m - 1, cuoi ? new Date(nam, m, 0).getDate() : d);
      if (x.getMonth() === m - 1) ra.push(x);
    };
    const R1 = /ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})(?:\s*(?:năm|\/)\s*(\d{4}))?/g;
    const R2 = /(?<![\d.,\/])(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{4}|\d{2}))?(?![.,]?\d)/g;
    const R3 = /(đầu|giữa|cuối)?\s*(?:tháng|(?<![\p{L}\d])t)\s*(\d{1,2})(?:\s*(?:năm|\/)\s*(\d{4}))?(?!\d)/gu;
    for (const m of t.matchAll(R1)) them(m[1], m[2], m[3]);
    t = t.replace(R1, " ");
    for (const m of t.matchAll(R2)) them(m[1], m[2], m[3]);
    t = t.replace(R2, " ");
    for (const m of t.matchAll(R3)) { const moc = { "đầu": 10, "giữa": 20 }[m[1] || ""]; them(moc || 1, m[2], m[3], !moc); }
    return ra.length ? new Date(Math.max(...ra)) : null;
  }
  const ngayVN = (d) => `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;

  // Công thức S1–S7 của Head of FIN (sheet AR Risk / Methodology)
  function signals(c, reason) {
    const J = c.total, Q = c.overdue, U = c.lwTotal || 0, V = c.lwOver || 0, W = c.lw61 || 0;
    return [
      +(U > 0 && J - U > 0.5 && Q - V >= -0.5), +(Q > 0.5 && Math.abs(Q - V) < 1),
      +(c.b[3] + c.b[4] + c.b[5] - W > 0.5), +(!reason.trim()), +(!String(c.term || "").trim()),
      +(J > 0.5 && Math.abs(Q - J) < 1),
      +(!!c.hen && c.hen < ASOF && Q > 0.5 && Q - V >= -0.5),
    ];
  }
  function trendOf(c) {
    const s = c.s, Q = c.overdue, U = c.lwTotal || 0, V = c.lwOver || 0, t = [];
    if (Q < 0.5) t.push("Chưa đến hạn thanh toán.");
    if (U < 0.5) t.push("Khách mới xuất hiện trong bảng tuần này.");
    if (s[0]) t.push("Khách phát sinh lô mới nhưng nợ cũ không giảm.");
    if (s[1]) t.push("Quá hạn đứng yên so với tuần trước, chưa thu được đồng nào.");
    if (s[6]) t.push(`Đã quá ngày hẹn trả (${ngayVN(c.hen)}) mà quá hạn chưa giảm.`);
    if (U > 0.5 && Q > 0.5 && Q < V - 0.5) t.push(`Đã thu bớt ${fmt.format(Math.round(V - Q))} so với tuần trước.`);
    if (s[2]) t.push("Nợ đang già đi, tiền chuyển sang nhóm tuổi cao hơn.");
    if (s[5]) t.push("Toàn bộ dư nợ đã quá hạn.");
    if (c.b[5] > 0.5) t.push(`Có ${fmt.format(Math.round(c.b[5]))} nằm ở nhóm trên 120 ngày.`);
    if (s[4]) t.push("Chưa có điều khoản thanh toán nên chưa xác định chắc chắn được quá hạn.");
    if (s[3]) t.push("CHƯA CÓ GIẢI TRÌNH TỪ SALES.");
    return t.join(" ");
  }

  function tierOf(c) {
    const reason = c.reason.startsWith(AUTO) ? "" : c.reason;
    // WR-11 PA1: Sales chưa ghi lý do tuần này thì xét nhóm (pháp lý, cấn trừ, ngày hẹn) theo lý do tuần trước
    const hieuLuc = reason.trim() ? reason : c.reasonLw.startsWith(AUTO) ? "" : c.reasonLw;
    const low = hieuLuc.toLowerCase();
    c.hen = ngayHen(hieuLuc) || c.pdFile || null;
    c.s = signals(c, reason);
    const sig = c.s[0] + c.s[1] + c.s[2] + c.s[6];
    let t;
    if (c.overdue < 0.5) t = "0·Chưa đến hạn";
    else if (KW_X.some((k) => low.includes(k))) t = "X·Không phải rủi ro tín dụng";
    else if (KW_N4.some((k) => low.includes(k))) t = "N4·Nợ xấu–sự kiện pháp lý";
    else if (c.b[5] > 0.5 || sig >= 2) t = "N3·Nợ xấu";
    else if (sig >= 1) t = "N2·Cần theo dõi";
    else t = "N1·Quá hạn bình thường";
    c.tier = t; c.tierKey = t.split("·")[0];
    c.ceo = c.overdue >= 100e6 && !reason.trim();
    c.trend = trendOf(c);
    c.action = c.tierKey === "N4" ? "Chuyển hồ sơ thu hồi. Rà soát chứng từ gốc để xét trích dự phòng theo TT48/2019."
      : c.tierKey === "N3" ? (c.b[5] > 0.5 ? "Dừng cấp dịch vụ mới. Chốt công văn đòi nợ có xác nhận gửi. Kế toán trưởng ký xác nhận phân nhóm."
        : "Dừng cấp dịch vụ mới. Yêu cầu cam kết NGÀY cụ thể bằng văn bản trong 7 ngày.") : "";
    // WR-11 (user 27/10): không hiện câu "[TỰ ĐỘNG] …" nữa
    c.reasonShown = reason || (c.reasonLw && !c.reasonLw.startsWith(AUTO) ? "Chưa có lý do tuần này · tuần trước: " + c.reasonLw
      : "Chưa có lý do tuần này");
    c.isAuto = !reason;
    // WR-13: mức rủi ro theo tiêu chí FIN — cùng luật wr/danhgia.py (ô trong file là công thức, có thể chưa có giá trị)
    const n = c.maxQh != null ? c.maxQh : c.overdue > 0.5 ? [0, 30, 60, 90, 120, 121][c.oldest] : 0;
    let m = n > 60 ? 3 : n > 30 ? 2 : 1;
    if (c.dg[0] === "Có" || (n > 0 && c.dg[2] === "Có")) m = 3;
    if (n > 0 && c.dg[1] === "Không") m = Math.max(m, 2);
    c.mucFin = m; c.mucTen = MUC_FIN[m - 1];
  }

  // ---------------- trang (tab) và địa chỉ ----------------
  // #sales?s=Trần Văn Hiếu&c=KH001&t=...&b=2&q=... — mỗi lần đổi trang / bộ lọc ghi vào đây
  function readHash() {
    const h = location.hash.slice(1), i = h.indexOf("?");
    const pg = decodeURIComponent(i < 0 ? h : h.slice(0, i)), qs = i < 0 ? "" : h.slice(i + 1);
    PAGE = PAGES.includes(pg) ? pg : "tong-quan";
    const p = new URLSearchParams(qs || "");
    F.sales = p.get("s") || ""; F.cust = p.get("c") || ""; F.tier = p.get("t") || "";
    F.bucket = p.has("b") ? +p.get("b") : -1; F.q = p.get("q") || "";
  }
  function writeHash() {
    const p = new URLSearchParams();
    if (F.sales) p.set("s", F.sales); if (F.cust) p.set("c", F.cust); if (F.tier) p.set("t", F.tier);
    if (F.bucket >= 0) p.set("b", F.bucket); if (F.q) p.set("q", F.q);
    const h = "#" + PAGE + (p.toString() ? "?" + p.toString() : "");
    if (location.hash === h) { syncControls(); render(); }
    else location.hash = h;              // hashchange sẽ vẽ lại
  }
  function go(page, patch) { if (page) PAGE = page; Object.assign(F, patch || {}); writeHash(); }

  // thanh tab: viên tím trượt theo chuột, đứng ở trang đang mở
  function navReady() {
    const nav = $("#jump"), pill = $("#pill"), links = [...nav.querySelectorAll("a")];
    const put = (a) => {
      if (!a) return;
      pill.style.left = a.offsetLeft + "px"; pill.style.width = a.offsetWidth + "px"; pill.style.opacity = 1;
      links.forEach((l) => l.classList.toggle("on", l === a));
    };
    const active = () => links.find((a) => a.dataset.p === PAGE);
    links.forEach((a) => {
      a.addEventListener("mouseenter", () => put(a));
      a.addEventListener("click", (e) => { e.preventDefault(); go(a.dataset.p); });
    });
    nav.addEventListener("mouseleave", () => put(active()));
    navReady.put = () => put(active());
    put(active());
  }

  // ---------------- bộ lọc chung ----------------
  function setupFilters() {
    const sales = new Set(), tiers = new Set();
    D.cust.forEach((c) => { salesOf(c).forEach((s) => sales.add(s)); tiers.add(c.tier); });
    opt($("#gSales"), [...sales].sort().map((s) => [s, s]), "Mọi Sales");
    opt($("#gTier"), [...tiers].sort().map((t) => [t, t]), "Mọi nhóm rủi ro");
    opt($("#gBucket"), BUCKET_VI.map((b, i) => [i, b]), "Mọi tuổi nợ");
    $("#gSales").onchange = (e) => setF({ sales: e.target.value });
    $("#gTier").onchange = (e) => setF({ tier: e.target.value });
    $("#gBucket").onchange = (e) => setF({ bucket: e.target.value === "" ? -1 : +e.target.value });
    let go_q = 0;
    $("#gSearch").oninput = (e) => { clearTimeout(go_q); go_q = setTimeout(() => setF({ q: e.target.value.trim().toLowerCase() }), 250); };
  }
  function opt(sel, pairs, first) { sel.innerHTML = `<option value="">${first}</option>` + pairs.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join(""); }
  function setF(p) { $("#tip").hidden = true; go(null, p); }
  function syncControls() {
    $("#gSales").value = F.sales; $("#gTier").value = F.tier; $("#gBucket").value = F.bucket < 0 ? "" : F.bucket;
    if ($("#gSearch").value.trim().toLowerCase() !== F.q) $("#gSearch").value = F.q;
    const ch = [];
    if (F.cust) ch.push(["cust", "Khách", (D.cust.find((c) => c.code === F.cust) || {}).name || F.cust]);
    if (F.sales) ch.push(["sales", "Sales", F.sales]);
    if (F.tier) ch.push(["tier", "Nhóm", F.tier]);
    if (F.bucket >= 0) ch.push(["bucket", "Tuổi nợ", BUCKET_VI[F.bucket]]);
    $("#chips").innerHTML = ch.map(([k, l, v]) => `<span class="chip" data-c="${k}" title="Bỏ lọc này"><b>${l}:</b> ${esc(v)} ✕</span>`).join("") +
      (ch.length > 1 ? `<span class="chip clear" data-c="*">Bỏ hết</span>` : "");
    $("#chips").querySelectorAll(".chip").forEach((c) => (c.onclick = () => c.dataset.c === "*"
      ? setF({ sales: "", cust: "", tier: "", bucket: -1, q: "" }) : setF({ [c.dataset.c]: c.dataset.c === "bucket" ? -1 : "" })));
  }
  const salesOf = (c) => {
    const s = new Set(c.sales.split(",").map((x) => x.trim()).filter(Boolean));
    c.kids.forEach((k) => k.sales && s.add(k.sales));
    return [...s];
  };
  const cf = () => D.cust.filter((c) =>
    (!F.sales || salesOf(c).includes(F.sales)) && (!F.tier || c.tier === F.tier) &&
    (F.bucket < 0 || c.b[F.bucket] > 0.5) && (!F.cust || c.code === F.cust) &&
    (!F.q || [c.name, c.code, c.job, c.kids.map((k) => k.job).join(" ")].join(" ").toLowerCase().includes(F.q)));

  // ---------------- khung thẻ: tiêu đề + chọn kiểu xem ----------------
  function card(id, title, sub, types, draw) {
    const el = $("#" + id);
    const cur = VIEW[id] && types.some((t) => t[0] === VIEW[id]) ? VIEW[id] : types[0][0];
    el.innerHTML = `<div class="card-h"><div class="grow"><h3>${title}</h3>${sub ? `<p class="sub">${sub}</p>` : ""}</div>
      ${types.length > 1 ? `<div class="seg">${types.map(([k, t]) => `<button data-v="${k}" class="${k === cur ? "act" : ""}">${t}</button>`).join("")}</div>` : ""}</div><div class="vz"></div>`;
    el.querySelectorAll(".seg button").forEach((b) => (b.onclick = () => { VIEW[id] = b.dataset.v; card(id, title, sub, types, draw); }));
    draw(cur, el.querySelector(".vz"));
  }

  // ---------------- vẽ ----------------
  function render() {
    const m = D.period.match(/(\d{2}\/\d{2}\/\d{4})\s*-\s*(\d{2}\/\d{2}\/\d{4})/);
    $("#periodChip b").textContent = m ? `${m[1].slice(0, 5)} – ${m[2]}` : D.period.replace(/^At\s*/, "");
    $("#srcNote").textContent = "Nguồn: " + D.file;
    const all = document.body.classList.contains("printing");   // in / lưu PDF: vẽ đủ mọi trang
    document.querySelectorAll(".page").forEach((p) => p.classList.toggle("on", all || p.dataset.page === PAGE));
    if (navReady.put) navReady.put();
    const L = cf(), on = (p) => all || PAGE === p;
    if (on("tong-quan")) { hero(); tiles(); aging(L); trend(); watch(L); salesMini(L); contracts(); }
    if (on("khach")) {
      const c = F.cust && D.cust.find((x) => x.code === F.cust);
      $("#khList").hidden = !!c; $("#khDetail").hidden = !c;
      crumb("#khCrumb", c ? [["Khách hàng", () => setF({ cust: "" })], [c.name]] : null);
      if (c) custPage(c); else khach(L);
    }
    if (on("sales")) {
      $("#saList").hidden = !!F.sales; $("#saDetail").hidden = !F.sales;
      crumb("#saCrumb", F.sales ? [["Sales", () => setF({ sales: "" })], [F.sales]] : null);
      if (F.sales) salesPage(F.sales, L); else { bySales(L); salesAge(L); salesTbl(L); }
    }
    if (on("rui-ro")) { relNote("#rrCrumb"); risk(L); tierMix(L); }
    if (on("tra")) vendors();
    if (on("tien")) { relNote("#ttCrumb"); cash(); }
  }

  // dòng điều hướng trên trang chi tiết: Khách hàng › TÊN   ← Tất cả khách
  function crumb(sel, parts) {
    const el = $(sel);
    if (!parts) { el.innerHTML = ""; el.hidden = true; return; }
    el.hidden = false;
    el.innerHTML = `<button class="back" type="button">← ${esc(parts[0][0])}: tất cả</button>` +
      parts.map(([t, f], i) => (f ? `<a href="#" data-i="${i}">${esc(t)}</a>` : `<b>${esc(t)}</b>`)).join(`<span class="sep">›</span>`);
    el.querySelector(".back").onclick = parts[0][1];
    el.querySelectorAll("a[data-i]").forEach((a) => (a.onclick = (e) => { e.preventDefault(); parts[+a.dataset.i][1](); }));
  }
  // trên tab khác mà đang lọc theo một khách / một Sales: nói rõ đang xem phần liên quan
  function relNote(sel) {
    const el = $(sel), bits = [];
    if (F.cust) bits.push(`khách <b>${esc((D.cust.find((c) => c.code === F.cust) || {}).name || F.cust)}</b>`);
    if (F.sales) bits.push(`Sales <b>${esc(F.sales)}</b>`);
    el.hidden = !bits.length;
    el.innerHTML = bits.length ? `<span class="rel">Chỉ hiện phần liên quan tới ${bits.join(" và ")}</span><button class="back" type="button">Bỏ lọc</button>` : "";
    if (bits.length) el.querySelector(".back").onclick = () => setF({ cust: "", sales: "" });
  }

  // tăng là tốt (+1), là xấu (-1) hay chỉ là thay đổi (0) — quyết định màu mũi tên
  const Y_NGHIA = { tong_phai_thu: -1, qua_han: -1, no_xau: -1, da_thu: 1, tong_phai_tra: 0, da_tra_vendor: 0, dong_tien_rong: 1, so_khach: 0 };
  function hero() {
    const n = D.now, l = D.last, d = (k) => (typeof n[k] === "number" && typeof l[k] === "number" ? n[k] - l[k] : null);
    const s = [];
    const dt = d("tong_phai_thu"), dq = d("qua_han");
    s.push(`Khách đang nợ <b>${short(n0(n.tong_phai_thu))}</b>` + (dt !== null ? `, <span class="${dt > 0 ? "neg" : "pos"}">${dt > 0 ? "tăng" : "giảm"} ${short(Math.abs(dt))}</span> so với tuần trước` : ""));
    s.push(`quá hạn <b>${short(n0(n.qua_han))}</b>` + (dq !== null ? ` (<span class="${dq > 0 ? "neg" : "pos"}">${dq > 0 ? "▲" : "▼"} ${short(Math.abs(dq))}</span>)` : ""));
    s.push(`nợ xấu <b class="neg">${short(n0(n.no_xau))}</b> ở ${n.so_khach_no_xau ?? 0} khách, ${n.ceo_can_xem ?? 0} khách cần CEO xem`);
    const cash = `Tuần này thu <b class="pos">${short(n0(n.da_thu))}</b>` + (typeof n.da_tra_vendor === "number" ? `, trả vendor <b>${short(n.da_tra_vendor)}</b> — dòng tiền ròng <b class="${n0(n.dong_tien_rong) >= 0 ? "pos" : "neg"}">${short(n0(n.dong_tien_rong))}</b>.` : ".");
    $("#hero").innerHTML = `<div class="hero-t">${s.join("; ")}.</div><div class="hero-s">${cash}</div>`;
  }

  // đường nhỏ trong ô số: giá trị các tuần của chỉ số đó
  function spark(k) {
    const H = D.history.map((h) => h[{ tong_phai_thu: "Total AR", qua_han: "Overdue", no_xau: "Bad debt", da_thu: "Collected",
      tong_phai_tra: "Total AP", da_tra_vendor: "Paid to vendors", dong_tien_rong: "Net cash" }[k]]).filter((v) => typeof v === "number");
    if (H.length < 2) return "";
    const mn = Math.min(...H), mx = Math.max(...H), W = 90, Hh = 30;
    const pts = H.map((v, i) => `${(i * W) / (H.length - 1)},${Hh - 3 - ((v - mn) / (mx - mn || 1)) * (Hh - 6)}`).join(" ");
    return `<svg class="spark" viewBox="0 0 ${W} ${Hh}" width="${W}" height="${Hh}"><polyline points="${pts}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  }

  function delta(k) {
    const a = D.now[k], b = D.last[k];
    if (typeof a !== "number" || typeof b !== "number" || !b) return "chưa có số tuần trước";
    const d = a - b, up = d >= 0, y = Y_NGHIA[k] ?? 0;
    const cls = y === 0 ? "neu" : (up === (y > 0) ? "pos" : "neg");
    return `<span class="${cls}">${up ? "▲" : "▼"} ${short(Math.abs(d))}</span> so với tuần trước`;
  }
  function tiles() {
    const n = D.now, tot = n0(n.tong_phai_thu);
    const T = [
      ["Tổng phải thu", n.tong_phai_thu, delta("tong_phai_thu"), ""],
      ["Quá hạn", n.qua_han, `${tot ? Math.round((n0(n.qua_han) / tot) * 100) : 0}% tổng phải thu · ` + delta("qua_han"), "bad"],
      ["Nợ xấu (N3 + N4)", n.no_xau, `${n.so_khach_no_xau ?? 0} khách · ${n.ceo_can_xem ?? 0} khách CEO cần xem · ` + delta("no_xau"), "bad"],
      ["Thu được tuần này", n.da_thu, delta("da_thu"), "cash"],
      ["Tổng phải trả vendor", n.tong_phai_tra, delta("tong_phai_tra"), "ap"],
      ["Đã trả vendor tuần này", n.da_tra_vendor, `nợ vendor mới ${short(n0(n.no_vendor_moi))}`, "ap"],
      ["Dòng tiền ròng tuần này", n.dong_tien_rong, "tiền về (NH + TM) − tiền trả vendor", n0(n.dong_tien_rong) >= 0 ? "cash" : "bad"],
      ["Số khách còn nợ", n.so_khach, `${n.so_vendor ?? "—"} vendor còn phải trả`, ""],
    ];
    const KEY = ["tong_phai_thu", "qua_han", "no_xau", "da_thu", "tong_phai_tra", "da_tra_vendor", "dong_tien_rong", "so_khach"];
    $("#tiles").innerHTML = T.map(([l, v, d, c], i) =>
      `<div class="tile ${c}"><div class="l">${l}</div>${spark(KEY[i])}<div class="v">${typeof v === "number" && Math.abs(v) < 1e4 ? fmt.format(v) : money(v)}</div><div class="d">${d}</div></div>`).join("");
  }

  // Tên ngắn cho nhãn biểu đồ: bỏ phần pháp lý / ngành nghề chung ở đầu tên ("CÔNG TY TNHH THƯƠNG MẠI ...")
  // để còn phần phân biệt khách. Tooltip và bảng vẫn hiện tên đầy đủ.
  const DAU_TEN = /^(CÔNG TY|CTY|TỔNG CÔNG TY|TNHH|CỔ PHẦN|CP|TRÁCH NHIỆM HỮU HẠN|MỘT THÀNH VIÊN|MTV|HAI THÀNH VIÊN TRỞ LÊN|THƯƠNG MẠI|TM|XUẤT NHẬP KHẨU|XNK|DỊCH VỤ|DV|ĐẦU TƯ|SẢN XUẤT|SX|VÀ|&|TẬP ĐOÀN|PHÁT TRIỂN|CÔNG NGHỆ|LLC|JSC|OOO)[\s,.-]+/i;
  const DUOI_TEN = /[\s,.-]+(VIỆT NAM|\(VIỆT NAM\)|VN|CO\.?,?\s*LTD\.?|COMPANY LIMITED|LIMITED|LTD\.?|PTE\.? LTD\.?|LLC|INC\.?|JSC|CORP\.?|GMBH|SP\. Z O\.O\.)$/i;
  function tenNgan(n) {
    let s = String(n || "").trim(), t;
    while ((t = s.replace(DAU_TEN, "")) !== s && t.length >= 2) s = t;
    while ((t = s.replace(DUOI_TEN, "")) !== s && t.length >= 2) s = t;
    return s || String(n || "");
  }
  const gon = (n, max) => { const s = tenNgan(n); return s.length > max ? s.slice(0, max - 1) + "…" : s; };

  // ai đang nợ: 5 khách lớn nhất của một nhóm, để đưa vào tooltip
  const topWho = (list, val, n = 5) => list.filter((c) => val(c) > 0.5).sort((a, b) => val(b) - val(a)).slice(0, n)
    .map((c) => [c.name, val(c)]);
  const tipBody = (title, total, of, who, extra) =>
    `<h4>${esc(title)}</h4><div class="tr"><span>Số tiền</span><b>${money(total)}</b></div>` +
    (of ? `<div class="tr"><span>Tỷ trọng</span><b>${pct(total, of)}</b></div>` : "") + (extra || "") +
    (who && who.length ? `<div class="mu">Ai đang nợ nhiều nhất</div>` + who.map(([n, v]) => `<div class="tr"><span>${esc(gon(n, 34))}</span><b>${short(v)}</b></div>`).join("") : "");

  function aging(L) {
    const v = BUCKETS.map((_, i) => L.reduce((a, c) => a + c.b[i], 0)), tot = v.reduce((a, b) => a + b, 0);
    const items = v.map((x, i) => ({ k: "b" + i, label: BUCKET_VI[i], value: x, color: AGE_COLORS[i],
      tip: tipBody(BUCKET_VI[i], x, tot, topWho(L, (c) => c.b[i])), click: () => setF({ bucket: F.bucket === i ? -1 : i }) }));
    card("cAging", "Tuổi nợ phải thu", "Bấm một nhóm tuổi để lọc cả trang", [["donut", "Donut"], ["bar", "Cột"], ["table", "Bảng"]],
      (t, el) => t === "donut" ? donut(el, items, "Tổng phải thu", tot) : t === "bar" ? hbar(el, items) : table(el,
        [["Nhóm tuổi", ""], ["Số tiền", "n"], ["Tỷ trọng", "n"]], items.map((x) => [x.label, money(x.value), pct(x.value, tot)])));
  }

  function trend() {
    const H = D.history.filter((h) => typeof h["Total AR"] === "number");
    const S = [["Total AR", NEU, "Tổng phải thu"], ["Overdue", NEG, "Quá hạn"], ["Collected", POS, "Thu được"]];
    card("cTrend", "Xu hướng các tuần", H.length < 2 ? "Mới có 1 tuần — từ tuần sau thành đường" : "Trỏ vào một tuần để xem số",
      [["line", "Đường"], ["col", "Cột"], ["table", "Bảng"]], (t, el) => {
        if (t === "table") return table(el, [["Tuần", ""], ...S.map((s) => [s[2], "n"])], H.map((h) => [h.Week, ...S.map((s) => money(h[s[0]]))]));
        lineCol(el, H, S, H.length < 2 ? "col" : t);
      });
  }

  function khach(L) {
    const top = [...L].sort((a, b) => b.total - a.total);
    card("cKhach", `Khách nợ nhiều nhất (${L.length} khách · ${short(L.reduce((a, c) => a + c.total, 0))})`,
      "Cột chia theo tuổi nợ. Bấm một khách để xem job và hoá đơn.", [["stack", "Cột chồng"], ["donut", "Donut"], ["table", "Bảng"]], (t, el) => {
        if (t === "table") return custTable(el, L);
        if (t === "donut") {
          const tot = top.reduce((a, c) => a + c.total, 0);
          const head = top.slice(0, 7), rest = top.slice(7);
          const items = head.map((c, i) => ({ k: c.code, label: c.name, value: c.total, color: CAT[i], tip: custTip(c), click: () => pick(c) }));
          if (rest.length) items.push({ k: "khac", label: `Khác (${rest.length} khách)`, value: rest.reduce((a, c) => a + c.total, 0), color: OTHER,
            tip: tipBody(`Khác (${rest.length} khách)`, rest.reduce((a, c) => a + c.total, 0), tot, topWho(rest, (c) => c.total)) });
          return donut(el, items, "Phải thu", tot);
        }
        stackBar(el, top.slice(0, 15).map((c) => ({ k: c.code, label: c.name, parts: c.b.map((v, i) => ({ value: v, color: AGE_COLORS[i] })),
          total: c.total, tip: custTip(c), click: () => pick(c) })), BUCKET_VI, AGE_COLORS);
      });
  }
  const custTip = (c) => `<h4>${esc(c.name)}</h4>
    <div class="tr"><span>Tổng nợ</span><b>${money(c.total)}</b></div><div class="tr"><span>Quá hạn</span><b>${money(c.overdue)}</b></div>
    <div class="tr"><span>Nhóm</span><b>${esc(c.tier)}</b></div><div class="tr"><span>Sales</span><b>${esc(salesOf(c).join(", ") || "—")}</b></div>
    ${c.kids.filter((k) => k.job).length ? `<div class="mu">${c.kids.filter((k) => k.job).length} job đang nợ · bấm để xem</div>` : ""}`;

  // bấm một khách ở bất kỳ đâu → trang chi tiết khách; bấm một Sales → trang của Sales đó
  const pick = (c) => c && go("khach", { cust: c.code });
  const pickSales = (s) => s && s !== "Chưa có Sales" && go("sales", { sales: s, cust: "" });

  // phần của một Sales trong một khách: các job họ phụ trách; khách không có job tính cho Sales của khách
  function portion(c, s) {
    const chu = c.sales.split(",")[0].trim();
    if (!c.kids.length) return chu === s ? { amount: c.total, b: [...c.b], jobs: [] } : null;
    const jobs = c.kids.filter((k) => (k.sales || chu) === s);
    if (!jobs.length) return null;
    return { amount: jobs.reduce((a, k) => a + k.amount, 0), b: BUCKETS.map((_, i) => jobs.reduce((a, k) => a + (k.b ? k.b[i] : 0), 0)), jobs };
  }
  function salesStats(L) {
    const m = {};
    L.forEach((c) => {
      const chu = c.sales.split(",")[0].trim() || "Chưa có Sales";
      const names = c.kids.length ? [...new Set(c.kids.map((k) => k.sales || chu))] : [chu];
      names.forEach((s) => {
        const p = portion(c, s === "Chưa có Sales" ? "" : s) || { amount: c.total, b: [...c.b], jobs: [] };
        const o = (m[s] = m[s] || { s, amount: 0, b: [0, 0, 0, 0, 0, 0], cust: [], jobs: 0, bad: 0 });
        o.amount += p.amount; p.b.forEach((v, i) => (o.b[i] += v)); o.cust.push([c, p.amount]); o.jobs += p.jobs.filter((k) => k.job).length;
        if (c.tierKey === "N3" || c.tierKey === "N4") o.bad += p.b.slice(1).reduce((a, v) => a + v, 0);
      });
    });
    return Object.values(m).sort((a, b) => b.amount - a.amount);
  }

  const nhomCua = (b) => b.map((v, i) => (v > 0.5 ? `<span class="agechip" style="--c:${AGE_COLORS[i]}">${BUCKETS[i]} · ${short(v)}</span>` : "")).join("");
  const DC = { "Khớp": "ok", "ĐỎ": "bad", "VÀNG": "warn", "CAM": "warn", "TÍM": "neu", "Khó đòi": "mute" };
  const dcBadge = (t) => { if (!t) return ""; const k = t.split(" ·")[0]; return `<span class="dc dc-${DC[k] || "mute"}" title="${esc(t)}">Đối chiếu: ${esc(k)}</span>`; };

  // ---------- trang chi tiết một khách ----------
  function custPage(c) {
    const pay = D.payers.find((p) => String(p.code).toUpperCase() === c.code.toUpperCase());
    $("#kdHead").innerHTML = `<div><h2>${esc(c.name)}</h2><p class="sub">${esc(c.code)} · ${esc(c.term || "chưa có term")} · ${badge(c)}
      ${mucBadge(c)} ${c.ceo ? '<span class="ceo">CEO cần xem</span>' : ""} ${dcBadge(c.dc)}${(c.alerts || []).map((a) => `<span class="hd ${a.muc}" title="${esc(a.note)}">${esc(a.loai)}</span>`).join("")}</p><p class="reason">${reasonCell(c)}</p>
      ${c.action ? `<p class="act">Việc cần làm: ${esc(c.action)}</p>` : ""}
      ${c.bySales ? `<p class="sub">Job theo Sales: ${esc(c.bySales)}</p>` : ""}</div>`;
    const dW = c.lwTotal ? c.total - c.lwTotal : null;
    $("#kdTiles").innerHTML = [
      ["Tổng nợ", money(c.total), dW === null ? "khách mới tuần này" : `<span class="${dW > 0 ? "neg" : "pos"}">${dW > 0 ? "▲" : "▼"} ${short(Math.abs(dW))}</span> so với tuần trước`, ""],
      ["Quá hạn", money(c.overdue), `${pct(c.overdue, c.total)} tổng nợ`, c.overdue > 0.5 ? "bad" : "cash"],
      ["Nợ già nhất", c.overdue > 0.5 ? BUCKET_VI[c.oldest] : "Chưa đến hạn", c.b[5] > 0.5 ? `${short(c.b[5])} trên 120 ngày` : "", c.oldest >= 3 ? "bad" : ""],
      ["Trả trong tuần", pay ? money(pay.total) : "—", pay ? `ngân hàng + tiền mặt ${short(pay.bank + pay.cash)}` : "chưa thấy tiền về tuần này", "cash"],
    ].map(([l, v, d, k]) => `<div class="tile ${k}"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join("");
    const items = c.b.map((v, i) => ({ k: "b" + i, label: BUCKET_VI[i], value: v, color: AGE_COLORS[i], tip: tipBody(BUCKET_VI[i], v, c.total) }));
    card("cKdAge", "Tuổi nợ của khách", "", [["donut", "Donut"], ["bar", "Cột"]], (t, el) => t === "donut" ? donut(el, items, "Tổng nợ", c.total) : hbar(el, items));
    const ss = {}; c.kids.forEach((k) => { const s = k.sales || c.sales.split(",")[0].trim() || "Chưa có Sales"; ss[s] = (ss[s] || 0) + k.amount; });
    if (!c.kids.length) ss[c.sales.split(",")[0].trim() || "Chưa có Sales"] = c.total;
    const si = Object.entries(ss).sort((a, b) => b[1] - a[1]).map(([s, v], i) => ({ k: "s" + i, label: s, value: v, color: CAT[i % CAT.length],
      tip: tipBody(s, v, c.total) + `<div class="mu">Bấm để mở trang của ${esc(s)}</div>`, click: () => pickSales(s) }));
    card("cKdSales", "Ai đang phụ trách", "Bấm một Sales để mở trang của người đó", [["donut", "Donut"], ["bar", "Cột"]],
      (t, el) => t === "donut" ? donut(el, si, "Job", si.reduce((a, x) => a + x.value, 0)) : hbar(el, si));
    card("cKdJobs", `Job đang nợ (${c.kids.length})`, c.dc && !c.dc.startsWith("Khớp") ? esc(c.dc) : "Lấy từ hoá đơn chưa thanh toán trên Misa", [["table", "Bảng"]], (t, el) => {
      if (!c.kids.length) return (el.innerHTML = `<p class="empty">Chưa có job — Sales theo báo cáo tuần trước: ${esc(c.sales || "—")}</p>`);
      el.innerHTML = `<div class="table-wrap"><table class="tbl small"><tr><th>Job · hoá đơn</th><th>Sales</th><th>Trạng thái</th><th>Nhóm tuổi</th><th class="n">Số tiền</th></tr>` +
        c.kids.map((k) => `<tr><td>${esc(k.label)}</td><td>${k.sales ? `<a href="#" class="lnk" data-s="${esc(k.sales)}">${esc(k.sales)}</a>` : "—"}</td><td>${esc(k.status || "")}</td><td>${nhomCua(k.b || [])}</td><td class="n">${money(k.amount)}</td></tr>`).join("") +
        `<tr><td colspan="4"><b>Cộng các job</b></td><td class="n"><b>${money(c.kids.reduce((a, k) => a + k.amount, 0))}</b></td></tr></table></div>`;
      el.querySelectorAll("a.lnk").forEach((a) => (a.onclick = (e) => { e.preventDefault(); pickSales(a.dataset.s); }));
    });
    card("cKdInv", `Hoá đơn còn nợ (${c.inv.length})`, "Theo sổ Misa sau khi trừ tiền đã thu", [["table", "Bảng"]], (t, el) => {
      el.innerHTML = `<div class="table-wrap"><table class="tbl small"><tr><th>Hoá đơn</th><th>Hạn</th><th class="n">Quá hạn</th><th class="n">Số tiền</th></tr>` +
        c.inv.map((v) => `<tr><td>${esc(v.no)}<br><small>${esc(v.date || "")}</small></td><td>${esc(v.due || "—")}</td><td class="n">${v.days ? v.days + " ngày" : "—"}</td><td class="n">${money(v.amount)}</td></tr>`).join("") +
        `<tr><td colspan="3"><b>Tổng</b></td><td class="n"><b>${money(c.total)}</b></td></tr></table></div>`;
    });
    const SIG = ["Nợ tăng mà quá hạn không giảm", "Quá hạn đứng yên so với tuần trước", "Nợ trên 60 ngày tăng",
                 "Chưa có giải trình của Sales", "Chưa có điều khoản thanh toán", "Toàn bộ dư nợ đã quá hạn",
                 "Lỡ hẹn: đã qua ngày Sales hẹn mà quá hạn không giảm"];
    card("cKdSig", "Dấu hiệu rủi ro", c.trend ? esc(c.trend) : "", [["list", "Danh sách"]], (t, el) => {
      el.innerHTML = c.s ? `<ul class="sig">${SIG.map((x, i) => `<li class="${c.s[i] ? "on" : ""}"><b>S${i + 1}</b> ${x}</li>`).join("")}</ul>` : `<p class="empty">Chưa có số AR Risk cho khách này.</p>`;
    });
  }

  // ---------- trang riêng một Sales ----------
  function salesPage(s, L) {
    const st = salesStats(L.filter((c) => salesOf(c).includes(s) || (!c.kids.length && c.sales.split(",")[0].trim() === s))).find((o) => o.s === s)
      || { s, amount: 0, b: [0, 0, 0, 0, 0, 0], cust: [], jobs: 0, bad: 0 };
    const over = st.b.slice(1).reduce((a, v) => a + v, 0);
    $("#sdHead").innerHTML = `<div><h2>${esc(s)}</h2><p class="sub">Chỉ phần job của Sales này. Bấm một khách để mở trang khách; các tab Nợ xấu, Dòng tiền cũng chỉ còn khách của ${esc(s)}.</p></div>`;
    $("#sdTiles").innerHTML = [
      ["Đang phụ trách", money(st.amount), `${pct(st.amount, D.cust.reduce((a, c) => a + c.total, 0))} tổng phải thu`, ""],
      ["Quá hạn", money(over), `${pct(over, st.amount)} phần phụ trách`, over > 0.5 ? "bad" : "cash"],
      ["Nợ xấu (N3 + N4)", money(st.bad), "phần quá hạn ở khách nợ xấu", st.bad > 0.5 ? "bad" : ""],
      ["Khách · job", `${st.cust.length} · ${st.jobs}`, "số khách và số job đang nợ", "ap"],
    ].map(([l, v, d, k]) => `<div class="tile ${k}"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join("");
    const rowsC = st.cust.sort((a, b) => b[1] - a[1]);
    card("cSdCust", `Khách của ${esc(s)} (${rowsC.length})`, "Cột chia theo tuổi nợ phần của Sales này. Bấm để mở trang khách.", [["stack", "Cột chồng"], ["table", "Bảng"]], (t, el) => {
      const R = rowsC.map(([c]) => ({ c, p: portion(c, s) || { amount: c.total, b: c.b, jobs: [] } }));
      if (t === "table") return table(el, [["Khách hàng", ""], ["Phần phụ trách", "n"], ["Quá hạn", "n"], ["Nhóm", ""]],
        R.map(({ c, p }) => [c.name, money(p.amount), money(p.b.slice(1).reduce((a, v) => a + v, 0)), c.tier]), (i) => pick(R[i].c));
      stackBar(el, R.slice(0, 15).map(({ c, p }) => ({ k: c.code, label: c.name, parts: p.b.map((v, i) => ({ value: v, color: AGE_COLORS[i] })), total: p.amount,
        tip: custTip(c), click: () => pick(c) })), BUCKET_VI, AGE_COLORS);
    });
    const ai = st.b.map((v, i) => ({ k: "b" + i, label: BUCKET_VI[i], value: v, color: AGE_COLORS[i], tip: tipBody(BUCKET_VI[i], v, st.amount) }));
    card("cSdAge", "Tuổi nợ phần phụ trách", "", [["donut", "Donut"], ["bar", "Cột"]], (t, el) => t === "donut" ? donut(el, ai, "Phụ trách", st.amount) : hbar(el, ai));
    const jobs = [];
    rowsC.forEach(([c]) => { const p = portion(c, s); (p ? p.jobs : []).forEach((k) => jobs.push([c, k])); });
    jobs.sort((a, b) => b[1].amount - a[1].amount);
    card("cSdJobs", `Job đang nợ (${jobs.length})`, "Bấm một dòng để mở trang khách", [["table", "Bảng"]], (t, el) =>
      jobs.length ? tableHtml(el, ["Job · hoá đơn", "Khách hàng", "Trạng thái", "Nhóm tuổi", "Số tiền"],
        jobs.map(([c, k]) => [esc(k.label), esc(c.name), esc(k.status || ""), nhomCua(k.b || []), money(k.amount)]), (i) => pick(jobs[i][0]))
        : (el.innerHTML = `<p class="empty">Không có job nào ghi tên ${esc(s)}.</p>`));
    const B = rowsC.map(([c]) => c).filter((c) => c.tierKey === "N3" || c.tierKey === "N4");
    card("cSdRisk", `Khách nợ xấu của ${esc(s)} (${B.length})`, "Lý do / nhận xét và việc cần làm", [["table", "Bảng"]], (t, el) =>
      B.length ? tableHtml(el, ["Khách hàng", "Nhóm", "Quá hạn", "Lý do / nhận xét", "Việc cần làm"],
        B.map((c) => [`<b>${esc(c.name)}</b>${c.ceo ? '<br><span class="ceo">CEO cần xem</span>' : ""}`, badge(c), money(c.overdue), reasonCell(c), `<small>${esc(c.action)}</small>`]), (i) => pick(B[i]))
        : (el.innerHTML = `<p class="empty">Không có khách nợ xấu.</p>`));
  }

  // ---------- tổng quan: cần chú ý + Sales ----------
  function watch(L) {
    const W = [...L].filter((c) => c.overdue > 0.5).sort((a, b) => b.overdue - a.overdue).slice(0, 8);
    card("cWatch", "Cần chú ý: quá hạn nhiều nhất", "Màu theo nhóm rủi ro. Bấm để mở trang khách.", [["bar", "Cột"], ["table", "Bảng"]], (t, el) => {
      if (!W.length) return (el.innerHTML = `<p class="empty">Không có khách quá hạn.</p>`);
      if (t === "table") return table(el, [["Khách hàng", ""], ["Quá hạn", "n"], ["Nhóm", ""]], W.map((c) => [c.name, money(c.overdue), c.tier]), (i) => pick(W[i]));
      hbar(el, W.map((c) => ({ k: c.code, label: c.name, value: c.overdue, color: TIER_COLOR[c.tierKey] || NEU, tip: custTip(c), click: () => pick(c) })));
    });
  }
  function contracts() {
    const A = (D.contracts || []).filter((a) => !F.sales || a.sales.includes(F.sales));
    const dem = (f) => A.filter(f).length;
    const sub = A.length ? `${dem((a) => a.muc === "do")} hết hạn còn nợ · ${dem((a) => a.loai.startsWith("Sắp"))} sắp hết hạn · ${dem((a) => a.loai.startsWith("Chưa có credit"))} chưa có term · ${dem((a) => a.loai.includes("vừa đổi"))} term vừa đổi. Bấm để mở trang khách.` : "";
    card("cContracts", "Hợp đồng & credit term", sub, [["table", "Bảng"]], (t, el) => {
      if (!D.contracts) return (el.innerHTML = `<p class="empty">File báo cáo này chưa có sheet Contracts.</p>`);
      if (!A.length) return (el.innerHTML = `<p class="empty">Không có cảnh báo hợp đồng / credit term tuần này.</p>`);
      tableHtml(el, ["Cảnh báo", "Khách hàng", "Sales", "Số tiền", "Hết hạn", "Còn lại", "Ghi chú"],
        A.map((a) => [`<span class="hd ${a.muc}">${esc(a.loai)}</span>`, esc(a.name), esc(a.sales || "—"), a.balance > 0.5 ? money(a.balance) : "—",
                      esc(a.han || "—"), a.left === null || a.left === undefined || a.left === "" ? "—" : esc(a.left + " ngày"), esc(a.note)]),
        (i) => pick(D.cust.find((c) => c.code.toUpperCase() === A[i].code.toUpperCase())));
    });
  }
  function salesMini(L) {
    const S = salesStats(L), tot = S.reduce((a, o) => a + o.amount, 0);
    card("cSalesMini", "Theo Sales", "Bấm một Sales để mở trang riêng", [["bar", "Cột"], ["donut", "Donut"]], (t, el) => {
      const items = S.map((o, i) => ({ k: "m" + i, label: o.s, value: o.amount, color: t === "bar" ? NEU : (i < 7 ? CAT[i] : OTHER),
        tip: tipBody(o.s, o.amount, tot, o.cust.sort((a, b) => b[1] - a[1]).slice(0, 5).map(([c, v]) => [c.name, v])), click: () => pickSales(o.s) }));
      t === "bar" ? hbar(el, items) : donut(el, items, "Phải thu", tot);
    });
  }

  function custTable(el, L) {
    const list = [...L].sort((a, b) => b.overdue - a.overdue || b.total - a.total);
    let h = `<div class="table-wrap"><table class="tbl"><tr><th>Khách hàng</th><th>Sales</th><th class="n">Tổng nợ</th><th class="n">Quá hạn</th><th>Nợ già nhất</th><th>Nhóm rủi ro</th><th>Lý do / nhận xét</th></tr>`;
    list.forEach((c) => (h += `<tr class="row" data-c="${esc(c.code)}"><td><b>${esc(c.name)}</b><br><small>${esc(c.code)} · ${esc(c.term || "chưa có term")}</small>${c.ceo ? '<br><span class="ceo">CEO cần xem</span>' : ""}</td>
      <td>${esc(salesOf(c).join(", ") || "—")}</td><td class="n">${money(c.total)}</td><td class="n">${money(c.overdue)}</td>
      <td>${c.overdue > 0.5 ? BUCKET_VI[c.oldest] : "—"}</td><td>${badge(c)}</td><td>${reasonCell(c)}</td></tr>`));
    el.innerHTML = h + "</table></div>";
    el.querySelectorAll("tr.row").forEach((tr) => (tr.onclick = () => pick(D.cust.find((c) => c.code === tr.dataset.c))));
  }

  function bySales(L) {
    const S = salesStats(L), tot = S.reduce((a, o) => a + o.amount, 0);
    const items = S.map((o, i) => ({ k: "s" + i, label: o.s, value: o.amount, color: i < 7 ? CAT[i] : OTHER,
      tip: tipBody(o.s, o.amount, tot, o.cust.sort((a, b) => b[1] - a[1]).slice(0, 5).map(([c, v]) => [c.name, v])) + `<div class="mu">Bấm để mở trang riêng</div>`,
      click: () => pickSales(o.s) }));
    card("cSales", "Ai đang phụ trách bao nhiêu", "Tính theo job; khách chưa có job tính cho Sales của khách", [["bar", "Cột"], ["donut", "Donut"]],
      (t, el) => t === "bar" ? hbar(el, items.map((x) => ({ ...x, color: NEU }))) : donut(el, items, "Phải thu", tot));
  }
  function salesAge(L) {
    const S = salesStats(L);
    card("cSalesAge", "Tuổi nợ theo Sales", "Bấm một Sales để mở trang riêng", [["stack", "Cột chồng"]], (t, el) =>
      stackBar(el, S.map((o) => ({ k: o.s, label: o.s, parts: o.b.map((v, i) => ({ value: v, color: AGE_COLORS[i] })), total: o.amount,
        tip: tipBody(o.s, o.amount, 0, null, o.b.map((v, i) => (v > 0.5 ? `<div class="tr"><span>${BUCKET_VI[i]}</span><b>${short(v)}</b></div>` : "")).join("")),
        click: () => pickSales(o.s) })), BUCKET_VI, AGE_COLORS));
  }
  function salesTbl(L) {
    const S = salesStats(L), tot = S.reduce((a, o) => a + o.amount, 0);
    card("cSalesTbl", "Bảng Sales", "Bấm một dòng để mở trang riêng", [["table", "Bảng"]], (t, el) =>
      table(el, [["Sales", ""], ["Phụ trách", "n"], ["Tỷ trọng", "n"], ["Quá hạn", "n"], ["Nợ xấu", "n"], ["Số khách", "n"], ["Số job", "n"]],
        S.map((o) => [o.s, money(o.amount), pct(o.amount, tot), money(o.b.slice(1).reduce((a, v) => a + v, 0)), money(o.bad), o.cust.length, o.jobs]),
        (i) => pickSales(S[i].s)));
  }

  const mucBadge = (c) => `<span class="muc muc-${c.mucFin}">${esc(c.mucTen)}</span>`;
  const badge = (c) => `<span class="badge b-${esc(c.tierKey)}">${esc(c.tier)}</span>`;
  const reasonCell = (c) => c.isAuto ? `<span class="auto">${esc(c.reasonShown)}</span>` : esc(c.reasonShown);

  function risk(L) {
    const B = L.filter((c) => c.tierKey === "N3" || c.tierKey === "N4").sort((a, b) => b.overdue - a.overdue);
    card("cRisk", `Nợ xấu N3 + N4 (${B.length} khách · ${short(B.reduce((a, c) => a + c.overdue, 0))} quá hạn)`, "Màu theo nhóm. Bấm một khách để xem chi tiết.",
      [["bar", "Cột"], ["table", "Bảng"]], (t, el) => {
        if (!B.length) return (el.innerHTML = `<p class="empty">Không có khách nợ xấu trong bộ lọc này.</p>`);
        if (t === "bar") return hbar(el, B.slice(0, 15).map((c) => ({ k: c.code, label: c.name, value: c.overdue, color: TIER_COLOR[c.tierKey],
          tip: custTip(c) + `<div class="mu">${esc(c.action || "")}</div>`, click: () => pick(c) })));
        let h = `<div class="table-wrap"><table class="tbl"><tr><th>Khách hàng</th><th>Mức độ (FIN)</th><th>Nhóm</th><th class="n">Quá hạn</th><th class="n">Trên 120 ngày</th><th>Lý do / nhận xét</th><th>Việc cần làm</th><th>Sales</th></tr>`;
        B.forEach((c) => (h += `<tr class="go" data-c="${esc(c.code)}"><td><b>${esc(c.name)}</b>${c.ceo ? '<br><span class="ceo">CEO cần xem</span>' : ""}</td><td>${mucBadge(c)}</td><td>${badge(c)}</td><td class="n">${money(c.overdue)}</td><td class="n">${money(c.b[5])}</td><td>${reasonCell(c)}</td><td><small>${esc(c.action)}</small></td><td>${esc(salesOf(c).join(", "))}</td></tr>`));
        el.innerHTML = h + "</table></div>";
        el.querySelectorAll("tr.go").forEach((tr) => (tr.onclick = () => pick(D.cust.find((c) => c.code === tr.dataset.c))));
      });
  }

  function tierMix(L) {
    const m = {};
    L.forEach((c) => { if (c.overdue > 0.5) (m[c.tier] = m[c.tier] || []).push(c); });
    const tot = L.reduce((a, c) => a + c.overdue, 0);
    const order = ["N4", "N3", "N2", "N1", "X"];
    const items = Object.entries(m).sort((a, b) => order.indexOf(a[0].split("·")[0]) - order.indexOf(b[0].split("·")[0])).map(([t, cs]) => {
      const v = cs.reduce((a, c) => a + c.overdue, 0);
      return { k: t, label: t, value: v, color: TIER_COLOR[t.split("·")[0]], tip: tipBody(t, v, tot, topWho(cs, (c) => c.overdue), `<div class="tr"><span>Số khách</span><b>${cs.length}</b></div>`),
        click: () => setF({ tier: F.tier === t ? "" : t }) };
    });
    card("cTier", "Quá hạn theo nhóm rủi ro", "Bấm một nhóm để lọc", [["donut", "Donut"], ["bar", "Cột"]],
      (t, el) => items.length ? (t === "donut" ? donut(el, items, "Quá hạn", tot) : hbar(el, items)) : (el.innerHTML = `<p class="empty">Không có nợ quá hạn.</p>`));
  }

  // Phải trả: FIN 09/10 — chỉ lấy số tổng, không phân tích
  function vendors() {
    const V = D.vendors.filter((v) => !F.q || (v.name + " " + v.code).toLowerCase().includes(F.q)).filter((v) => v.close > 0.5).sort((a, b) => b.close - a.close);
    const tot = V.reduce((a, v) => a + v.close, 0);
    const vt = (v) => `<h4>${esc(v.name)}</h4><div class="tr"><span>Còn phải trả</span><b>${money(v.close)}</b></div><div class="tr"><span>Tỷ trọng</span><b>${pct(v.close, tot)}</b></div>` +
      (v.paid > 0.5 ? `<div class="tr"><span>Đã trả tuần này</span><b>${money(v.paid)}</b></div>` : "");
    card("cVendor", `Mình đang nợ ai (${V.length} vendor · ${short(tot)})`, "15 vendor nợ nhiều nhất", [["bar", "Cột"], ["table", "Bảng"]], (t, el) => {
      if (!V.length) return (el.innerHTML = `<p class="empty">File tuần này chưa có số phải trả.</p>`);
      if (t === "bar") return hbar(el, V.slice(0, 15).map((v) => ({ k: v.code, label: v.name, value: v.close, color: NEU, tip: vt(v) })));
      table(el, [["Vendor", ""], ["Mã", ""], ["Còn phải trả", "n"], ["Tỷ trọng", "n"]], V.map((v) => [v.name, v.code, money(v.close), pct(v.close, tot)]));
    });
    const items = V.slice(0, 7).map((v, i) => ({ k: v.code, label: v.name, value: v.close, color: CAT[i], tip: vt(v) }));
    const rest = V.slice(7);
    if (rest.length) items.push({ k: "khac", label: `Khác (${rest.length} vendor)`, value: rest.reduce((a, v) => a + v.close, 0), color: OTHER,
      tip: tipBody(`Khác (${rest.length} vendor)`, rest.reduce((a, v) => a + v.close, 0), tot) });
    card("cVStatus", "Nợ vendor tập trung ở đâu", "7 vendor lớn nhất và phần còn lại", [["donut", "Donut"]],
      (t, el) => items.length ? donut(el, items, "Phải trả", tot) : (el.innerHTML = `<p class="empty">Chưa có số phải trả.</p>`));
  }

  function cash() {
    const n = D.now;
    const ch = [["Về ngân hàng", n.thu_ngan_hang], ["Tiền mặt", n.thu_tien_mat], ["Cấn trừ với phải trả", n.thu_can_tru], ["Tỷ giá, phí ngân hàng", n.thu_ty_gia_phi], ["Khác", n.thu_khac]]
      .filter((x) => n0(x[1]) > 0.5);
    const tot = ch.reduce((a, x) => a + n0(x[1]), 0);
    const MAU_KENH = { "Về ngân hàng": POS, "Tiền mặt": "#008300", "Cấn trừ với phải trả": NEU, "Tỷ giá, phí ngân hàng": "#f08a3c", "Khác": OTHER };
    const items = ch.map(([l, v], i) => ({ k: "c" + i, label: l, value: n0(v), color: MAU_KENH[l], tip: tipBody(l, n0(v), tot) }));
    card("cChannel", "Tiền thu từ khách theo kênh", `Tổng ${short(tot)}`, [["donut", "Donut"], ["bar", "Cột"], ["table", "Bảng"]],
      (t, el) => t === "donut" ? donut(el, items, "Đã thu", tot, true) : t === "bar" ? hbar(el, items, true) : table(el, [["Kênh", ""], ["Số tiền", "n"], ["Tỷ trọng", "n"]], items.map((x) => [x.label, money(x.value), pct(x.value, tot)])));
    const ma = new Set(cf().map((c) => c.code.toUpperCase()));
    const loc = F.cust || F.sales || F.tier || F.bucket >= 0 || F.q;
    const P = D.payers.filter((p) => !loc || ma.has(String(p.code).toUpperCase())).slice(0, 10);
    const cOf = (p) => D.cust.find((c) => c.code.toUpperCase() === String(p.code).toUpperCase());
    card("cPayers", "Khách trả nhiều nhất", "10 khách", [["bar", "Cột"], ["table", "Bảng"]], (t, el) => t === "bar"
      ? hbar(el, P.map((p) => ({ k: p.code, label: p.name, value: p.total, color: POS, click: cOf(p) ? () => pick(cOf(p)) : null, tip: `<h4>${esc(p.name)}</h4><div class="tr"><span>Ngân hàng + tiền mặt</span><b>${money(p.bank + p.cash)}</b></div><div class="tr"><span>Cấn trừ, tỷ giá, khác</span><b>${money(p.offset + p.fx + p.other)}</b></div>` })), true)
      : table(el, [["Khách hàng", ""], ["Tổng", "n"]], P.map((p) => [p.name, money(p.total)])));
    const V = D.vendors.filter((v) => v.paid > 0.5).sort((a, b) => b.paid - a.paid).slice(0, 10);
    card("cPaidV", "Đã trả vendor nhiều nhất", "10 vendor", [["bar", "Cột"], ["table", "Bảng"]], (t, el) => !V.length ? (el.innerHTML = `<p class="empty">Chưa có số.</p>`) : t === "bar"
      ? hbar(el, V.map((v) => ({ k: v.code, label: v.name, value: v.paid, color: NEU, tip: `<h4>${esc(v.name)}</h4><div class="tr"><span>Đã trả tuần này</span><b>${money(v.paid)}</b></div><div class="tr"><span>Còn phải trả</span><b>${money(v.close)}</b></div>` })), true)
      : table(el, [["Vendor", ""], ["Đã trả", "n"]], V.map((v) => [v.name, money(v.paid)])));
  }

  // ---------------- các kiểu biểu đồ (SVG tự vẽ, không thư viện) ----------------
  function donut(el, items, center, total, compact) {
    const R = 100, r = 62, C = 110;
    let a0 = -Math.PI / 2, g = "";
    const tot = items.reduce((a, x) => a + Math.max(0, x.value), 0) || 1;
    items.forEach((x) => {
      if (x.value <= 0.5) return;
      const a1 = a0 + (2 * Math.PI * x.value) / tot, gap = items.length > 1 ? 0.012 : 0;
      const p = (rr, a) => [C + rr * Math.cos(a), C + rr * Math.sin(a)];
      const [s, e] = [a0 + gap, Math.max(a0 + gap, a1 - gap)], big = e - s > Math.PI ? 1 : 0;
      const [x1, y1] = p(R, s), [x2, y2] = p(R, e), [x3, y3] = p(r, e), [x4, y4] = p(r, s);
      g += x.value >= tot - 0.5 ? `<circle cx="${C}" cy="${C}" r="${(R + r) / 2}" fill="none" stroke="${x.color}" stroke-width="${R - r}" data-k="${esc(x.k)}"/>`
        : `<path d="M${x1},${y1} A${R},${R} 0 ${big} 1 ${x2},${y2} L${x3},${y3} A${r},${r} 0 ${big} 0 ${x4},${y4} Z" fill="${x.color}" data-k="${esc(x.k)}"/>`;
      a0 = a1;
    });
    g += `<text x="${C}" y="${C - 4}" text-anchor="middle" style="font-size:12px">${esc(center)}</text><text x="${C}" y="${C + 16}" text-anchor="middle" style="font:700 17px var(--fh);fill:var(--ink)">${short(total)}</text>`;
    const lg = items.map((x) => `<div class="row" data-k="${esc(x.k)}"><span class="sw" style="background:${x.color}"></span><span>${esc(gon(x.label, 38))}</span><b class="n">${short(x.value)}</b><span class="n">${pct(x.value, tot)}</span></div>`).join("");
    el.innerHTML = `<div class="donut-wrap"><svg viewBox="0 0 220 220" width="220" height="220" role="img">${g}</svg><div class="lg">${lg}</div></div>`;
    wire(el, items);
  }

  function hbar(el, items, narrow) {
    const max = Math.max(...items.map((x) => Math.abs(x.value)), 1);
    const W = Math.max(360, el.clientWidth || (narrow ? 560 : 1100)), rowH = 30, R = 80, H = items.length * rowH + 8;
    const L = Math.min(300, Math.round(W * 0.34)), maxCh = Math.floor((L - 16) / 7.9);
    let g = "";
    items.forEach((x, i) => {
      const w = Math.max(6, ((W - L - R) * Math.abs(x.value)) / max), y = 4 + i * rowH;
      const lab = gon(x.label, maxCh);
      g += `<g data-k="${esc(x.k)}"><rect x="0" y="${y}" width="${W}" height="${rowH}" fill="transparent"/>
        <text x="${L - 10}" y="${y + 19}" text-anchor="end" style="fill:var(--ink-2);font-size:12px">${esc(lab)}</text>
        <path d="M${L},${y + 6} h${w - 4} a4,4 0 0 1 4,4 v10 a4,4 0 0 1 -4,4 h-${w - 4} z" fill="${x.color}"/>
        <text x="${L + w + 8}" y="${y + 19}" style="fill:var(--ink);font-size:12px">${short(x.value)}</text></g>`;
    });
    el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img">${g}</svg>`;
    wire(el, items);
  }

  function stackBar(el, rows, names, colors) {
    const max = Math.max(...rows.map((r) => r.total), 1);
    const W = Math.max(360, el.clientWidth || 1100), rowH = 30, R = 80, H = rows.length * rowH + 8;
    const L = Math.min(300, Math.round(W * 0.34)), maxCh = Math.floor((L - 16) / 7.9);
    let g = "";
    rows.forEach((r, i) => {
      const y = 4 + i * rowH; let x = L;
      g += `<g data-k="${esc(r.k)}"><rect x="0" y="${y}" width="${W}" height="${rowH}" fill="transparent"/><text x="${L - 10}" y="${y + 19}" text-anchor="end" style="fill:var(--ink-2);font-size:12px">${esc(gon(r.label, maxCh))}</text>`;
      r.parts.forEach((p) => { if (p.value <= 0.5) return; const w = ((W - L - R) * p.value) / max; g += `<rect x="${x}" y="${y + 6}" width="${Math.max(1, w - 2)}" height="18" rx="2" fill="${p.color}"/>`; x += w; });
      g += `<text x="${x + 8}" y="${y + 19}" style="fill:var(--ink);font-size:12px">${short(r.total)}</text></g>`;
    });
    el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img">${g}</svg>
      <div class="legend">${names.map((n, i) => `<span><i style="background:${colors[i]}"></i>${n}</span>`).join("")}</div>`;
    wire(el, rows);
  }

  function lineCol(el, H, S, kind) {
    const W = Math.max(360, el.clientWidth || 560), Hh = 240, L = 64, R = 30, T = 14, B = 28;
    const max = Math.max(...H.flatMap((h) => S.map((s) => n0(h[s[0]])))) * 1.1 || 1;
    const band = (W - L - R) / Math.max(H.length, 1);
    const x = (i) => L + band * (i + 0.5);
    const y = (v) => T + (Hh - T - B) * (1 - v / max);
    let g = "";
    for (let k = 0; k <= 4; k++) { const v = (max * k) / 4, yy = y(v); g += `<line x1="${L}" x2="${W - R}" y1="${yy}" y2="${yy}" stroke="var(--grid)"/><text x="${L - 8}" y="${yy + 4}" text-anchor="end">${short(v)}</text>`; }
    H.forEach((h, i) => (g += `<text x="${x(i)}" y="${Hh - 8}" text-anchor="middle">${esc(h.Week)}</text>`));
    if (kind === "col") {
      const bw = Math.min(18, (band - 12) / S.length);
      H.forEach((h, i) => S.forEach(([k, c], j) => { const v = n0(h[k]); const xx = x(i) - (bw * S.length) / 2 + j * bw; g += `<rect class="col" x="${xx + 1}" y="${y(v)}" width="${bw - 2}" height="${y(0) - y(v)}" rx="3" fill="${c}"/>`; }));
    } else {
      S.forEach(([k, c]) => {
        if (H.length > 1) g += `<polyline points="${H.map((h, i) => `${x(i)},${y(n0(h[k]))}`).join(" ")}" fill="none" stroke="${c}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
        H.forEach((h, i) => (g += `<circle cx="${x(i)}" cy="${y(n0(h[k]))}" r="5" fill="${c}" stroke="#fff" stroke-width="2"/>`));
      });
    }
    // vùng trỏ theo tuần: trỏ vào cột tuần nào thì hiện đủ số của tuần đó
    const items = H.map((h, i) => ({ k: "w" + i, tip: `<h4>${esc(h.Week)} · đến ${esc(h["Period end"])}</h4>` + S.map(([k, c, l]) => `<div class="tr"><span><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${c};margin-right:5px"></i>${l}</span><b>${money(h[k])}</b></div>`).join("") }));
    H.forEach((h, i) => (g += `<rect data-k="w${i}" x="${x(i) - band / 2}" y="${T}" width="${band}" height="${Hh - T - B}" fill="transparent"/>`));
    el.innerHTML = `<svg viewBox="0 0 ${W} ${Hh}" width="100%" role="img">${g}</svg><div class="legend">${S.map(([, c, l]) => `<span><i style="background:${c}"></i>${l}</span>`).join("")}</div>`;
    wire(el, items, true);
  }

  function table(el, cols, rows, onRow) {
    el.innerHTML = `<div class="table-wrap"><table class="tbl small"><tr>${cols.map(([c, cl]) => `<th class="${cl}">${esc(c)}</th>`).join("")}</tr>` +
      rows.map((r, i) => `<tr${onRow ? ` class="go" data-i="${i}"` : ""}>${r.map((v, j) => `<td class="${cols[j][1]}">${esc(v)}</td>`).join("")}</tr>`).join("") + `</table></div>`;
    if (onRow) el.querySelectorAll("tr.go").forEach((tr) => (tr.onclick = () => onRow(+tr.dataset.i)));
  }
  // như table() nhưng ô đã là HTML (badge, chip tuổi nợ); cột có tên "Số tiền" / "Quá hạn" canh phải
  function tableHtml(el, cols, rows, onRow) {
    const so = (c) => (/^(Số tiền|Quá hạn)$/.test(c) ? "n" : "");
    el.innerHTML = `<div class="table-wrap"><table class="tbl small"><tr>${cols.map((c) => `<th class="${so(c)}">${esc(c)}</th>`).join("")}</tr>` +
      rows.map((r, i) => `<tr${onRow ? ` class="go" data-i="${i}"` : ""}>${r.map((v, j) => `<td class="${so(cols[j])}">${v}</td>`).join("")}</tr>`).join("") + `</table></div>`;
    if (onRow) el.querySelectorAll("tr.go").forEach((tr) => (tr.onclick = () => onRow(+tr.dataset.i)));
  }

  // ---------------- trỏ chuột & bấm ----------------
  const tip = $("#tip");
  function wire(el, items, noDim) {
    const by = Object.fromEntries(items.map((x) => [String(x.k), x]));
    el.querySelectorAll("[data-k]").forEach((m) => {
      const x = by[m.dataset.k]; if (!x) return;
      m.addEventListener("mousemove", (e) => {
        tip.innerHTML = x.tip || ""; tip.hidden = !x.tip;
        const tw = tip.offsetWidth, left = e.clientX + 16 + tw > innerWidth ? e.clientX - tw - 16 : e.clientX + 16;
        tip.style.left = left + "px"; tip.style.top = Math.min(e.clientY + 14, innerHeight - tip.offsetHeight - 8) + "px";
        if (!noDim) { el.classList.add("hovering"); el.querySelectorAll(`[data-k="${CSS.escape(m.dataset.k)}"]`).forEach((n) => n.classList.add("hot")); }
      });
      m.addEventListener("mouseleave", () => { tip.hidden = true; el.classList.remove("hovering"); el.querySelectorAll(".hot").forEach((n) => n.classList.remove("hot")); });
      if (x.click) m.addEventListener("click", x.click);
    });
  }
})();
