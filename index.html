<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Energy Savings — Checklist Hub (หมวดหมู่) · v2.2</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    .shadow-soft{box-shadow:0 6px 20px rgba(0,0,0,.12)}
    .row-hover:hover{background:#0b122005}
    .sticky-head th{position:sticky;top:0;background:#fff}
    .toast{position:fixed;right:1rem;bottom:1rem;padding:.75rem 1rem;border-radius:.75rem;z-index:50}
    .active-cat{background:#0f172a;color:#e2e8f0}
    pre.code{background:#0b1220; color:#e2e8f0; padding:1rem; border-radius:.75rem; overflow:auto}
    details>summary{cursor:pointer}
  </style>
</head>
<body class="bg-slate-50 text-slate-800">
  <div class="max-w-7xl mx-auto p-4 md:p-8">
    <header class="mb-6">
      <h1 class="text-2xl md:text-3xl font-semibold">Energy Savings — Checklist Hub (หมวดหมู่)</h1>
      <p class="text-slate-600">จัดระเบียบเช็กลิสต์ตามหมวด (HVAC/M&E/Energy) • บันทึก/โหลดข้อมูลจาก Google Sheets • v2.2</p>
    </header>

    <section class="grid grid-cols-1 lg:grid-cols-[260px_1fr] gap-4">
      <!-- Sidebar: Categories -->
      <aside class="bg-white rounded-xl shadow-soft h-max sticky top-4">
        <div class="p-3 border-b border-slate-200">
          <label class="block text-sm font-medium">Google Apps Script Web App URL</label>
          <input id="webapp" class="mt-1 w-full rounded-lg border-slate-300" placeholder="วาง URL จาก Deploy > Web app" />
          <p class="text-xs text-slate-500 mt-1">ระบบจำใน localStorage</p>
        </div>
        <nav id="catList" class="p-2">
          <!-- populated by JS -->
        </nav>
      </aside>

      <!-- Main Panel -->
      <main>
        <!-- Controls -->
        <section class="mb-4 grid gap-3 md:grid-cols-2 lg:grid-cols-4">
          <div class="p-3 bg-white rounded-xl shadow-soft">
            <div class="text-sm text-slate-500">หมวดที่เลือก</div>
            <div class="text-lg font-semibold" id="catTitle">—</div>
            <div class="text-xs text-slate-500" id="catSheet">Sheet: —</div>
          </div>
          <div class="p-3 bg-white rounded-xl shadow-soft">
            <label class="block text-sm font-medium">ค้นหา</label>
            <input id="search" class="mt-1 w-full rounded-lg border-slate-300" placeholder="ค้นหา Task/Owner/Notes" />
          </div>
          <div class="p-3 bg-white rounded-xl shadow-soft flex items-end gap-2">
            <button id="btnLoad" class="px-4 py-2 rounded-lg bg-slate-800 text-white">โหลดจาก Google Sheet</button>
            <button id="btnSave" class="px-4 py-2 rounded-lg bg-emerald-600 text-white">บันทึกทั้งหมด</button>
            <button id="btnCSV" class="px-4 py-2 rounded-lg bg-sky-600 text-white">Export CSV</button>
          </div>
          <div class="p-3 bg-white rounded-xl shadow-soft flex items-end">
            <button id="btnTest" class="px-4 py-2 rounded-lg bg-indigo-600 text-white w-full" title="Run smoke tests">Run Tests</button>
          </div>
        </section>

        <!-- Progress -->
        <section class="mb-4 p-3 bg-white rounded-xl shadow-soft">
          <div class="flex items-center justify-between mb-2">
            <div class="text-sm font-medium">ความคืบหน้า</div>
            <div class="text-sm"><span id="doneCount">0</span>/<span id="totalCount">0</span> งาน Done</div>
          </div>
          <div class="w-full bg-slate-200 rounded-full h-3">
            <div id="bar" class="h-3 rounded-full bg-emerald-500" style="width:0%"></div>
          </div>
        </section>

        <!-- Table -->
        <section class="bg-white rounded-xl shadow-soft overflow-hidden">
          <div class="overflow-x-auto">
            <table class="w-full text-sm">
              <thead class="sticky-head">
                <tr class="bg-slate-100 border-b border-slate-200">
                  <th class="p-3 text-left">ID</th>
                  <th class="p-3 text-left w-[32rem]">Task</th>
                  <th class="p-3 text-left">Owner</th>
                  <th class="p-3 text-left">Due</th>
                  <th class="p-3 text-left">Status</th>
                  <th class="p-3 text-left">Evidence (Link)</th>
                  <th class="p-3 text-left">Notes</th>
                </tr>
              </thead>
              <tbody id="tbody"></tbody>
            </table>
          </div>
        </section>

        <footer class="text-xs text-slate-500 mt-4">
          <p>Tip: สถานะ ToDo/Doing/Done มีผลต่อแถบความคืบหน้า · ใช้ Evidence วางลิงก์ Google Drive/รูป/ไฟล์ · สลับหมวดที่แถบซ้าย</p>
        </footer>
      </main>
    </section>
  </div>

  <div id="toast" class="toast bg-slate-900 text-white hidden"></div>

  <script>
  // ====== Config & Constants ======
  const DEFAULT_STATUS = ["ToDo", "Doing", "Done"]; // ต้องตรงกับ Apps Script

  // ====== Helper: Date parsing/formatting ======
  function pad2(n){ return String(n).padStart(2,'0'); }
  function toDateInputValue(v){
    if(!v) return '';
    // If already in YYYY-MM-DD
    if(typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    // Google returns ISO when JSON.stringify(Date)
    if(typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) return v.slice(0,10);
    // dd/mm/yyyy
    if(typeof v === 'string' && /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.test(v)){
      const [_,d,m,y] = v.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      return `${y}-${pad2(m)}-${pad2(d)}`;
    }
    // Serial number from Sheets (rare if we stringify dates, but handle anyway)
    if(typeof v === 'number'){
      const epoch = new Date(Date.UTC(1899,11,30));
      const dt = new Date(epoch.getTime()+v*86400000);
      return `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth()+1)}-${pad2(dt.getUTCDate())}`;
    }
    // Date object
    if(v instanceof Date) {
      const y=v.getFullYear(), m=pad2(v.getMonth()+1), d=pad2(v.getDate());
      return `${y}-${m}-${d}`;
    }
    return '';
  }

  // ====== Category & Task Definitions ======
  const TASKS_SITE = [
    {id:"S01", text:"บิลค่าไฟ 12–24 เดือนล่าสุด (kWh, kW, Ft, PF, Total)"},
    {id:"S02", text:"โหลดโปรไฟล์ 15-min (3–12 เดือน ครอบคลุมช่วง peak) หรืออย่างน้อย AM/PM curve"},
    {id:"S03", text:"ประเภทอัตราไฟ/สัญญา TOU/DED + Peak hours"},
    {id:"S04", text:"PF และค่าปรับ (ค่าปรับ/ส่วนลด PF)"},
    {id:"S05", text:"อุณหภูมิภายนอก (CDD/HDD) สำหรับโครงการที่ HVAC เด่น"},
    {id:"S06", text:"แผนผังไฟ/Single Line และจุดวัดสำคัญ"},
    {id:"S07", text:"รายการอุปกรณ์หลัก (Chiller, AHU, Pump, Compressor, Boiler, Lighting, Motors) + ปี/ยี่ห้อ/การควบคุม"},
    {id:"S08", text:"ชั่วโมงเดินเครื่องจริง ต่อวัน/สัปดาห์ (ไม่ใช่ตามคู่มือ)"},
    {id:"S09", text:"Setpoint/Control Logic ปัจจุบัน (เช่น Chilled water, Static pressure, VFD min/max)"},
    {id:"S10", text:"ปัญหาหน้างาน: จุดอั้น, Hot/Cold complaints, Maintenance backlog"},
    {id:"S11", text:"วัฒนธรรมการใช้งาน: เปิด–ปิดก่อน/หลังเวลางาน, Override บ่อยไหม"},
    {id:"S12", text:"ข้อจำกัดงานติดตั้ง: พื้นที่/การหยุดการผลิต/เสียง/ฝุ่น/ความปลอดภัย"},
    {id:"S13", text:"มาตรฐาน/ข้อกำหนด (ESG/KPI ฝ่ายบริหาร, BMS policy)"},
    {id:"S14", text:"ต้นทุนพลังงานอื่น (น้ำ, ไอน้ำ, NG/ดีเซล) หากเกี่ยวข้อง"},
    {id:"S15", text:"โอกาส ECM ด่วน (Quick wins): setpoint tune, schedule, VFD min, AHU coil clean"},
    {id:"S16", text:"ECM หลัก 2–3 ทางเลือก (Good/Better/Best) พร้อมช่วง CapEx/ผลประหยัด"},
    {id:"S17", text:"Assumptions/Exclusions—เขียนชัด เพื่อเทียบ bid ได้แฟร์"},
    {id:"S18", text:"Baseline model: เลือก IPMVP Option (A/B/C/D), ตัวแปรอิสระ (Temp/Prod)"},
    {id:"S19", text:"คุณภาพข้อมูล: R², CV(RMSE) เบื้องต้นถ้าทำ regression"},
    {id:"S20", text:"M&V Plan One-Pager: boundary, จุดวัด, ความถี่, routine/non-routine adj."},
    {id:"S21", text:"Risk register: ผลกระทบ×โอกาสเกิด + mitigation"},
    {id:"S22", text:"TCO 5 ปี: CapEx + OpEx − Savings + Risk (Downtime×฿/hr)"},
    {id:"S23", text:"Sensitivity: ค่าไฟ +5/10/15% ส่งผล Payback อย่างไร"},
    {id:"S24", text:"Comparison Matrix (Bid Equalization): ทำตารางเปรียบเทียบทุกเจ้าให้เทียบได้จริง"},
    {id:"S25", text:"Executive Summary 1 หน้า: Why change/now/us + KPI | ROI ≥X% | Payback ≤Y mo"},
    {id:"S26", text:"MAP 30 วัน (Mutual Action Plan): ขั้นสรุปจนออก PO"},
    {id:"S27", text:"เอกสารแนบ: รูป, บิล, SLD, Log บำรุงรักษา, Data export links"},
    {id:"S28", text:"เจ้าของงาน (Owner) ต่อฟิลด์: ตั้งคนรับผิดชอบและกำหนดวันครบกำหนด"},
    {id:"S29", text:"สถานะ–หลักฐาน: ลิงก์ Google Drive/รูป/ไฟล์แนบ สำหรับ audit trail"},
    {id:"S30", text:"สรุป PoC (ถ้าจำเป็น): ขอบเขต, ระยะเวลา, เกณฑ์ผ่าน–ไม่ผ่าน"}
  ];

  // Domain Checklists (ย่อ 8–12 จุดต่อหมวด)
  const TASKS_CHILLER = [
    {id:"C01", text:"ระบุชนิด (Centrifugal/Screw/Absorption) + ปี/รุ่น/ขนาด"},
    {id:"C02", text:"COP/kWTR ปัจจุบัน @ ภาระ 100/75/50%"},
    {id:"C03", text:"Chilled Water Supply/Return Setpoint + ΔT"},
    {id:"C04", text:"Condenser Water Temp & Approach (CW in/out, ambient/WB)"},
    {id:"C05", text:"คอมเพรสเซอร์/ปั๊มน้ำ/VFD มี/ไม่มี และ min/max speed"},
    {id:"C06", text:"Sequencing/Auto staging/Anti‑short cycling logic"},
    {id:"C07", text:"คุณภาพน้ำ/ตะกรัน/สเกล/บำรุงรักษาล่าสุด"},
    {id:"C08", text:"Energy Saving Ideas: reset CHWS, tower WB reset, VFD tuning"}
  ];
  const TASKS_PUMP = [
    {id:"P01", text:"ชนิดปั๊ม + BEP + duty/standby"},
    {id:"P02", text:"Throttling/Bypass อยู่ตรงไหนบ้าง (สูญเสีย)"},
    {id:"P03", text:"VFD/Soft‑starter + min/max freq"},
    {id:"P04", text:"Differential pressure setpoint + sensor location"},
    {id:"P05", text:"Balance/commissioning ล่าสุด"},
    {id:"P06", text:"Energy Ideas: DP reset, parallel control, impeller trim"}
  ];
  const TASKS_MOTOR = [
    {id:"M01", text:"Nameplate: kW, V, A, PF, Eff (IE2/IE3/IE4)"},
    {id:"M02", text:"โหลดเฉลี่ย (%) และ Running hours"},
    {id:"M03", text:"Power quality: THDi/THDv, อุณหภูมิเบื้องต้น"},
    {id:"M04", text:"VFD suitability/เดิมติด VFD หรือไม่"},
    {id:"M05", text:"Energy Ideas: High‑eff motor, VFD, soft‑start scheduling"}
  ];
  const TASKS_HEATPUMP = [
    {id:"H01", text:"ชนิด (Air/Water source) + Capacity"},
    {id:"H02", text:"Hot water setpoint + COP/SCOP"},
    {id:"H03", text:"Integration กับระบบน้ำร้อน/ย้อนน้ำกลับ"},
    {id:"H04", text:"Idea: heat recovery จากชิลเลอร์/คอนเดนซิ่ง"}
  ];
  const TASKS_AC = [
    {id:"A01", text:"ชนิด (DX/VRF/AHU+FCU) + ขนาด/ปี"},
    {id:"A02", text:"Thermostat/Setpoint/Deadband + Schedule"},
    {id:"A03", text:"Ventilation/OA rate + CO₂ control"},
    {id:"A04", text:"Idea: VAV/VSD fan, economizer/CO₂‑based control"}
  ];
  const TASKS_CT = [
    {id:"T01", text:"CW fan VFD + Basin level + Drift/bleed"},
    {id:"T02", text:"Approach/WB reset"},
    {id:"T03", text:"Idea: fan VFD tuning, fill/driver efficiency"}
  ];
  const TASKS_PV = [
    {id:"PV1", text:"Installed DC (kWp) / AC (kW) + DC/AC ratio"},
    {id:"PV2", text:"Inverter รุ่น/ประสิทธิภาพ/MPPT count"},
    {id:"PV3", text:"PR/Specific yield (kWh/kWp/yr)"},
    {id:"PV4", text:"Shading/Soiling/IV curve test plan"},
    {id:"PV5", text:"Idea: string re‑group, cleaning schedule, clipping mgmt"}
  ];
  const TASKS_BESS = [
    {id:"B1", text:"Rated kW/kWh, DoD, Cycle life"},
    {id:"B2", text:"EMS objectives: peak‑shave, TOU arbitrage, backup"},
    {id:"B3", text:"Safety: NFPA/IEC, room ventilation, isolation"},
    {id:"B4", text:"Idea: tariff‑aware scheduling, PV‑coupled charging"}
  ];
  const TASKS_ME_ELEC = [
    {id:"E1", text:"SLD update + Main meter list"},
    {id:"E2", text:"PF correction: capacitor health, detuning"},
    {id:"E3", text:"Harmonics: THDv/THDi จุดวิกฤต"},
    {id:"E4", text:"Idea: demand control, VT schedules, sensor QA"}
  ];
  const TASKS_ME_PIPING = [
    {id:"PP1", text:"Steam/Compressed Air leak survey"},
    {id:"PP2", text:"Pressure setpoint & drop across network"},
    {id:"PP3", text:"Heat recovery/Insulation audit"},
    {id:"PP4", text:"Idea: no‑load control, leak bounty, VSD compressor"}
  ];

  const CATEGORIES = [
    {id:'site',    name:'Site Checklist (30 จุด)',              sheet:'00_Site_Checklist_30', tasks:TASKS_SITE},
    {id:'chiller', name:'HVAC — Chiller',                        sheet:'HVAC_Chiller',        tasks:TASKS_CHILLER},
    {id:'pump',    name:'HVAC — Pump',                           sheet:'HVAC_Pump',           tasks:TASKS_PUMP},
    {id:'motor',   name:'Electrical — Motor',                    sheet:'ELEC_Motor',          tasks:TASKS_MOTOR},
    {id:'heatpump',name:'HVAC — Heat Pump',                      sheet:'HVAC_HeatPump',       tasks:TASKS_HEATPUMP},
    {id:'ac',      name:'HVAC — Air Conditioning',               sheet:'HVAC_AC',             tasks:TASKS_AC},
    {id:'ct',      name:'HVAC — Cooling Tower',                  sheet:'HVAC_CoolingTower',   tasks:TASKS_CT},
    {id:'pv',      name:'Energy — Solar PV',                     sheet:'ENE_SolarPV',         tasks:TASKS_PV},
    {id:'bess',    name:'Energy — BESS (Battery Storage)',       sheet:'ENE_BESS',            tasks:TASKS_BESS},
    {id:'mee',     name:'M&E — Electrical & Controls',           sheet:'ME_Electrical',       tasks:TASKS_ME_ELEC},
    {id:'mep',     name:'M&E — Piping/Steam/Compressed Air',     sheet:'ME_Piping',           tasks:TASKS_ME_PIPING}
  ];

  // ====== DOM refs ======
  const $ = sel => document.querySelector(sel);
  const tbody = $("#tbody");
  const totalCountEl = $("#totalCount");
  const doneCountEl = $("#doneCount");
  const bar = $("#bar");
  const toast = $("#toast");
  const urlInput = $("#webapp");
  const catList = $("#catList");
  const catTitle = $("#catTitle");
  const catSheet = $("#catSheet");
  const btnSave = $("#btnSave");

  // ====== State ======
  let CURRENT = null;              // current category object
  const STATE = {};                // sheetName -> array of records

  // Restore WebApp URL
  urlInput.value = localStorage.getItem("WEB_APP_URL") || "";
  urlInput.addEventListener("change", ()=>{ localStorage.setItem("WEB_APP_URL", urlInput.value.trim()); pingToast("บันทึก URL แล้ว ✓"); });

  function pingToast(msg, ok=true){
    toast.textContent = msg; toast.classList.remove("hidden");
    toast.style.background = ok?"#16a34a":"#b91c1c"; setTimeout(()=>toast.classList.add("hidden"), 2500);
  }

  // ====== Renderers ======
  function renderSidebar(){
    catList.innerHTML = CATEGORIES.map(c=>`<button data-id="${c.id}" class="w-full text-left px-3 py-2 rounded-lg mb-1 hover:bg-slate-100">${c.name}</button>`).join("");
    catList.addEventListener('click', e=>{
      const btn = e.target.closest('button[data-id]'); if(!btn) return;
      const cat = CATEGORIES.find(x=>x.id===btn.dataset.id);
      switchCategory(cat);
      [...catList.querySelectorAll('button')].forEach(b=>b.classList.remove('active-cat'));
      btn.classList.add('active-cat');
    });
  }

  function renderRows(tasks){
    tbody.innerHTML = tasks.map(r=>{
      return `<tr class="border-b border-slate-200 row-hover">
        <td class="p-3 align-top text-slate-500">${r.id}</td>
        <td class="p-3 align-top">${r.text}</td>
        <td class="p-2 align-top"><input class="w-36 md:w-40 rounded border-slate-300" data-k="Owner" data-id="${r.id}"/></td>
        <td class="p-2 align-top"><input type="date" class="w-36 md:w-40 rounded border-slate-300" data-k="Due_Date" data-id="${r.id}"/></td>
        <td class="p-2 align-top">
          <select class="w-28 md:w-32 rounded border-slate-300" data-k="Status" data-id="${r.id}">
            ${DEFAULT_STATUS.map(s=>`<option value="${s}">${s}</option>`).join("")}
          </select>
        </td>
        <td class="p-2 align-top"><input class="w-52 md:w-64 rounded border-slate-300" placeholder="https://..." data-k="Evidence" data-id="${r.id}"/></td>
        <td class="p-2 align-top"><input class="w-64 md:w-full rounded border-slate-300" data-k="Notes" data-id="${r.id}"/></td>
      </tr>`
    }).join("");
    updateProgress();
  }

  function getAllData(){
    const inputs = tbody.querySelectorAll("input, select");
    const map = {};
    (CURRENT?.tasks||[]).forEach(t=>{ map[t.id] = {Task_ID:t.id, Task:t.text, Owner:"", Due_Date:"", Status:"ToDo", Evidence:"", Notes:""}; });
    inputs.forEach(el=>{
      const id=el.getAttribute("data-id"), k=el.getAttribute("data-k"), v=(el.value||"").trim();
      if(k==='Due_Date') map[id][k] = toDateInputValue(v); else if(map[id]) map[id][k]=v;
    });
    return Object.values(map);
  }

  function applyRecords(records){
    if(!Array.isArray(records)) return; const map={}; records.forEach(r=>map[r.Task_ID]=r);
    (CURRENT?.tasks||[]).forEach(t=>{
      ["Owner","Due_Date","Status","Evidence","Notes"].forEach(k=>{
        const el = tbody.querySelector(`[data-id="${t.id}"][data-k="${k}"]`);
        if(!el) return; const val = map[t.id]?.[k];
        if(k==='Due_Date') el.value = toDateInputValue(val || '');
        else if(val!==undefined) el.value = val;
      });
    });
    updateProgress();
  }

  function updateProgress(){
    const inputs = Array.from(tbody.querySelectorAll('select[data-k="Status"]'));
    const total = inputs.length; const done = inputs.filter(el=>el.value==="Done").length;
    totalCountEl.textContent = total; doneCountEl.textContent = done;
    const pct = total? Math.round((done/total)*100) : 0; bar.style.width = pct+"%";
  }

  function switchCategory(cat){
    if(CURRENT){ STATE[CURRENT.sheet] = getAllData(); }
    CURRENT = cat; catTitle.textContent = cat.name; catSheet.textContent = `Sheet: ${cat.sheet}`;
    renderRows(cat.tasks); applyRecords(STATE[cat.sheet]);
  }

  // ====== Search ======
  const searchEl = document.querySelector('#search');
  searchEl.addEventListener('input', e=>{
    const q = e.target.value.toLowerCase();
    const base = (CURRENT?.tasks||[]);
    const rows = base.filter(t=>{
      const r = tbody.querySelector(`[data-id="${t.id}"]`);
      const owner = r?.closest('tr')?.querySelector('[data-k="Owner"]').value.toLowerCase()||"";
      const notes = r?.closest('tr')?.querySelector('[data-k="Notes"]').value.toLowerCase()||"";
      return t.id.toLowerCase().includes(q) || t.text.toLowerCase().includes(q) || owner.includes(q) || notes.includes(q);
    });
    renderRows(rows); applyRecords(STATE[CURRENT.sheet]);
  });

  // ====== Helpers ======
  function buildCSV(rows){
    if(!rows || rows.length===0) return '';
    const headers = Object.keys(rows[0]);
    return [headers.join(",")]
      .concat(rows.map(r=>headers.map(h=>`"${String(r[h]||"").replaceAll('"','""')}"`).join(",")))
      .join("\n");
  }

  // ====== Google Sheets IO ======
  async function saveAll(){
    const WEB_APP_URL = (localStorage.getItem("WEB_APP_URL")||urlInput.value||"").trim();
    if(!WEB_APP_URL){ pingToast("ยังไม่ได้ใส่ Web App URL", false); return; }
    STATE[CURRENT.sheet] = getAllData();
    const payload = {records: STATE[CURRENT.sheet], sheet: CURRENT.sheet};

    // UI: disable button + spinner
    btnSave.disabled = true; const old = btnSave.textContent; btnSave.textContent = 'กำลังบันทึก...'; btnSave.classList.add('opacity-70');

    try{
      const res = await fetch(WEB_APP_URL, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
      if(!res.ok) throw new Error("HTTP "+res.status);
      const data = await res.json().catch(()=>({}));
      const count = Array.isArray(data.updated) ? data.updated.length : payload.records.length;
      pingToast(`บันทึกสำเร็จ ✓ อัปเดต ${count} รายการ → ${CURRENT.sheet}`);
    }catch(err){
      try{
        await fetch(WEB_APP_URL, {method:"POST", mode:"no-cors", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
        pingToast(`ส่งข้อมูลแล้ว (no‑cors) → ${CURRENT.sheet} — เปิดชีตตรวจได้เลย`);
      }catch(e){ pingToast("ส่งไม่สำเร็จ: "+e.message, false); }
    }finally{
      btnSave.disabled = false; btnSave.textContent = old; btnSave.classList.remove('opacity-70');
    }
  }

  async function loadAll(){
    const WEB_APP_URL = (localStorage.getItem("WEB_APP_URL")||urlInput.value||"").trim();
    if(!WEB_APP_URL){ pingToast("ยังไม่ได้ใส่ Web App URL", false); return; }
    try{
      const url = WEB_APP_URL + `?sheet=${encodeURIComponent(CURRENT.sheet)}`;
      const res = await fetch(url);
      const data = await res.json();
      if(data && data.records){ STATE[CURRENT.sheet] = data.records; applyRecords(STATE[CURRENT.sheet]); pingToast(`โหลดข้อมูลสำเร็จ ✓ จาก ${CURRENT.sheet}`); }
      else pingToast("รูปแบบข้อมูลไม่คุ้น", false);
    }catch(err){ pingToast("โหลดไม่สำเร็จ (cors/สิทธิ์): "+err.message, false); }
  }

  document.querySelector('#btnSave').addEventListener('click', saveAll);
  document.querySelector('#btnLoad').addEventListener('click', loadAll);
  document.querySelector('#tbody').addEventListener('change', e=>{ if(e.target && e.target.matches('select[data-k="Status"]')) updateProgress(); });

  // Export CSV (เฉพาะหมวดที่เลือก)
  document.querySelector('#btnCSV').addEventListener('click', ()=>{
    STATE[CURRENT.sheet] = getAllData();
    const rows = STATE[CURRENT.sheet]||[]; if(rows.length===0){ pingToast('ไม่มีข้อมูล'); return; }
    const csv = buildCSV(rows);
    const blob = new Blob([csv], {type:'text/csv'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${CURRENT.sheet}.csv`; a.click();
  });

  // ====== Lightweight Test Harness ======
  function assert(name, cond){ (cond?console.log:console.error)(`${cond?'PASS':'FAIL'}: ${name}`); return !!cond; }
  function runTests(){
    const results = [];
    results.push(assert('Category list non-empty', CATEGORIES.length>0));

    // switch through all categories and ensure rows render
    CATEGORIES.forEach((c, i)=>{ switchCategory(c); results.push(assert(`Render ${c.id}`, tbody.querySelectorAll('tr').length === c.tasks.length)); });

    // back to default
    switchCategory(CATEGORIES[0]);
    results.push(assert('Default category set', !!CURRENT && CURRENT.id==='site'));

    // Progress logic
    const selects = Array.from(tbody.querySelectorAll('select[data-k="Status"]'));
    selects.forEach(s=>s.value='Done'); updateProgress();
    results.push(assert('Progress Done == Total', Number(doneCountEl.textContent)===Number(totalCountEl.textContent)));

    // CSV builder
    const sample = [{Task_ID:'T1', Task:'Demo', Owner:'A', Due_Date:'2025-09-13', Status:'Done', Evidence:'http://x', Notes:'ok'}];
    const csv = buildCSV(sample);
    results.push(assert('CSV header', csv.startsWith('Task_ID,Task,Owner,Due_Date,Status,Evidence,Notes')));
    results.push(assert('CSV 2 lines', csv.split('\n').length===2));

    // Payload shape
    const payload = {records:getAllData(), sheet: CURRENT.sheet};
    const r0 = payload.records[0]||{};
    results.push(assert('Payload shape ok', ['Task_ID','Task','Owner','Due_Date','Status','Evidence','Notes'].every(k=>k in r0)));

    const pass = results.every(Boolean);
    pingToast(pass?`Tests passed (${results.filter(Boolean).length}/${results.length})`:`Some tests failed (${results.filter(Boolean).length}/${results.length})`, pass);
  }
  document.querySelector('#btnTest').addEventListener('click', runTests);

  // ====== Init ======
  function init(){
    renderSidebar();
    switchCategory(CATEGORIES[0]);
    const firstBtn = catList.querySelector('button'); if(firstBtn) firstBtn.classList.add('active-cat');
  }
  init();
  </script>

  <!-- ================= Apps Script (Code.gs) =================
  แปะไว้เพื่ออ้างอิงเท่านั้น (คัดลอกไปวางใน Apps Script ของ Google Sheet)
  การปรับปรุง v2.2:
   • ตั้งค่า Data Validation ของคอลัมน์ Status ให้เป็น Dropdown (ToDo/Doing/Done)
   • ตั้ง Number Format ของ Due_Date เป็น yyyy-mm-dd
   • ปรับ doGet ให้ยิง Due_Date เป็นสตริง yyyy-mm-dd เพื่อเข้ากับ input[type=date]
  -->
  <details class="max-w-7xl mx-auto p-4 md:p-8">
    <summary class="font-medium">ดูโค้ด Apps Script (คลิกเพื่อแสดง)</summary>
    <pre class="code"><code>const SPREADSHEET_ID = 'YOUR_SHEET_ID';
const SHEET_NAME_DEFAULT = '00_Site_Checklist_30';
const STATUS_LIST = ['ToDo','Doing','Done']; // ต้องตรงกับหน้าเว็บ

function _getSheet(name){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaderAndFormat(sh){
  const header = ['Task_ID','Task','Owner','Due_Date','Status','Evidence','Notes'];
  if(sh.getLastRow()===0){ sh.getRange(1,1,1,header.length).setValues([header]); }
  // Format Due_Date (col 4)
  sh.getRange(2,4,Math.max(1000, sh.getMaxRows()-1),1).setNumberFormat('yyyy-mm-dd');
  // Data validation for Status (col 5)
  const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(STATUS_LIST, true)
              .setAllowInvalid(false).build();
  sh.getRange(2,5,Math.max(1000, sh.getMaxRows()-1),1).setDataValidation(rule);
}

function parseDateText(v){
  if(!v) return '';
  if(Object.prototype.toString.call(v)==='[object Date]') return v; // already Date
  if(typeof v==='number'){ // serial
    return new Date(1899,11,30+v);
  }
  if(typeof v==='string'){
    if(/^\d{4}-\d{2}-\d{2}$/.test(v)){
      const [y,m,d]=v.split('-').map(Number); return new Date(y,m-1,d);
    }
    const m=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/yyyy
    if(m){ const d=Number(m[1]), mo=Number(m[2]), y=Number(m[3]); return new Date(y,mo-1,d); }
  }
  return '';
}

function toYMD(v){
  if(!v) return '';
  let d=v;
  if(Object.prototype.toString.call(v)!=='[object Date]') d=parseDateText(v);
  if(!d) return '';
  const y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), da=('0'+d.getDate()).slice(-2);
  return `${y}-${m}-${da}`;
}

function doGet(e){
  const name = (e && e.parameter && e.parameter.sheet) || SHEET_NAME_DEFAULT;
  const sh = _getSheet(name);
  ensureHeaderAndFormat(sh);
  const values = sh.getDataRange().getValues();
  if(values.length < 2){
    return ContentService.createTextOutput(JSON.stringify({records:[]}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const headers = values[0];
  const rows = values.slice(1).map(r=>{
    const o = {}; headers.forEach((h,i)=>o[String(h||'')] = r[i]);
    // normalize date out
    o.Due_Date = toYMD(o.Due_Date);
    return o;
  });
  return ContentService.createTextOutput(JSON.stringify({records: rows}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const lock = LockService.getScriptLock(); lock.tryLock(20000);
  try{
    const body = e?.postData?.contents || '{}';
    const data = JSON.parse(body);
    const name = data.sheet || SHEET_NAME_DEFAULT;
    const sh = _getSheet(name);
    ensureHeaderAndFormat(sh);

    const header = ['Task_ID','Task','Owner','Due_Date','Status','Evidence','Notes'];

    // Build index by Task_ID
    const last = sh.getLastRow();
    const rng = last>1 ? sh.getRange(2,1,last-1,header.length).getValues() : [];
    const idx = {}; rng.forEach((r,i)=>{ idx[String(r[0])] = i+2; });

    const rows = Array.isArray(data.records) ? data.records : [];
    const out = [];
    rows.forEach(rec=>{
      // normalize inputs
      const due = parseDateText(rec.Due_Date);
      const status = String(rec.Status||'ToDo');
      const row = [
        rec.Task_ID||'', rec.Task||'', rec.Owner||'', due || '', status, rec.Evidence||'', rec.Notes||''
      ];
      const at = idx[String(rec.Task_ID)] || 0;
      if(at){ sh.getRange(at,1,1,row.length).setValues([row]); }
      else { sh.appendRow(row); idx[String(rec.Task_ID)] = sh.getLastRow(); }
      out.push(rec.Task_ID);
    });

    // re-apply formatting (in case of new rows)
    ensureHeaderAndFormat(sh);

    return ContentService.createTextOutput(JSON.stringify({ok:true, updated:out}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }finally{ try{lock.releaseLock()}catch(e){} }
}
</code></pre>
  </details>

</body>
</html>
