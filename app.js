// === Сводка ҚБМ — главный JS ===
const STORAGE_KEY = 'svodka_data_v1';
const SYNC_URL_KEY = 'svodka_sync_url';

// Структура листа "1-Мәліметтер": каждый параметр = 12 строк × 31 колонка дней
const PARAM_ROWS = {
  qn_tov:     5,  // Парктік өнім Qн (товарная), тн
  qn_fact:   17,  // Нақты мұнай Qм (фактическая), тн
  loss:      29,  // Технологические потери, тн
  sdacha:    41,  // Сдача Qм, тн
  park:      53,  // Остаток в парке, тн
  qzh_tot:   65,  // Общая жидкость, м³
  qzh_clean: 77,  // Жидкость без регенерации, м³
  dns_in:    89,  // ДНС вход, м³
  obv:       90,  // Обводнённость на ДНС, %
  par_all:  161,  // Пар всего (ППГ-3 + СПГУ + МПГУ), тн
  drop_w:   173,  // Сброс воды, м³
  bitum:    185,  // Битум, тн
  ozpv:     209,  // ОЗПВ опресн. вода, м³
};
// Лист "Показатели" — месячные суммы (колонка B = название, C..N = месяцы)
const SUMMARY_ROWS = {
  dob:4, sd:5, bit:6, zh:7, park:9,
  pg3:11, mpgu:12, spgu:13, par:14,
  zak_int:20, zak_ext_sum:27,
  vlz:31, gaz:35, ozpv:38,
};

const MONTHS = ['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек'];
const MONTH_NAMES_RU = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];

// === ХЕЛПЕРЫ ===
const fmt = (v, dec=0) => v==null||isNaN(v) ? '—' : Number(v).toLocaleString('ru-RU',{minimumFractionDigits:dec,maximumFractionDigits:dec});
const fmtK = (v) => v==null||isNaN(v) ? '—' : Math.abs(v)>=1e6 ? (v/1e6).toFixed(2)+' млн' : Math.abs(v)>=1e3 ? (v/1e3).toFixed(1)+' тыс' : fmt(v);

function $(id){return document.getElementById(id)}

// === ПАРСИНГ XLSX ===
async function parseXLSX(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, {type:'array', cellDates:true});
  const data = { daily:{}, monthly:{}, fund:{}, lastDate:null };

  // 1. Лист "1-Мәліметтер" — суточные данные
  const sName1 = wb.SheetNames.find(n=>n.includes('Мәліметтер')) || wb.SheetNames[0];
  const ws1 = wb.Sheets[sName1];
  const a1 = XLSX.utils.sheet_to_json(ws1, {header:1, defval:null, raw:true});
  // a1[r-1][c-1] = ячейка (r,c)
  for(const [key, startRow] of Object.entries(PARAM_ROWS)){
    data.daily[key] = {};
    for(let mi=0; mi<12; mi++){
      const row = a1[startRow-1+mi] || [];
      const days = [];
      for(let d=0; d<31; d++){
        const v = row[3+d]; // колонка D = индекс 3
        days.push(typeof v==='number' ? v : null);
      }
      data.daily[key][MONTHS[mi]] = days;
    }
  }

  // 2. Лист "Показатели" — месячные сводки
  const sName2 = wb.SheetNames.find(n=>n.trim().toLowerCase().startsWith('показатели'));
  if(sName2){
    const ws2 = wb.Sheets[sName2];
    const a2 = XLSX.utils.sheet_to_json(ws2, {header:1, defval:null, raw:true});
    for(const [key, r] of Object.entries(SUMMARY_ROWS)){
      const row = a2[r-1] || [];
      data.monthly[key] = [];
      for(let m=0; m<12; m++){
        const v = row[2+m]; // колонка C = индекс 2
        data.monthly[key].push(typeof v==='number' ? v : 0);
      }
    }
  }

  // 3. Лист "Сут.рапортАО КБМ" — фонд
  const sName3 = wb.SheetNames.find(n=>n.includes('рапорт'));
  if(sName3){
    const ws3 = wb.Sheets[sName3];
    const a3 = XLSX.utils.sheet_to_json(ws3, {header:1, defval:null, raw:true});
    const get = (r,c)=>{const row=a3[r-1]; return row ? row[c-1] : null;};
    data.fund = {
      flow: get(18,4), pump: get(18,6), total: get(18,8),
      inFlow: get(18,12), inPump: get(18,14),
      injSteamE: get(44,3), injWaterE: get(44,5), injTotalE: get(44,7),
      injSteamA: get(44,9), injWaterA: get(44,11), injTotalA: get(44,13),
    };
  }

  // 4. Определяем последнюю дату
  data.lastDate = findLastDate(data.daily);
  return data;
}

function findLastDate(daily){
  const now = new Date();
  const y = now.getFullYear();
  // Ищем последний месяц, у которого есть данные по qn_tov, и последний день в нём
  for(let mi=11; mi>=0; mi--){
    const days = daily.qn_tov[MONTHS[mi]] || [];
    for(let d=30; d>=0; d--){
      if(days[d]!=null && days[d]!==0){
        return { y, m:mi, d:d+1, monthName:MONTH_NAMES_RU[mi] };
      }
    }
  }
  return null;
}

// === РЕНДЕР ДАШБОРДА ===
function render(data){
  if(!data || !data.lastDate){
    $('lastDate').textContent = 'нет данных — загрузите XLSX';
    return;
  }
  const {y,m,d,monthName} = data.lastDate;
  $('lastDate').textContent = `${d} ${MONTH_NAMES_RU[m].toLowerCase()} ${y}`;

  // Главное: добыча Qн за сутки
  const qnArr = data.daily.qn_tov[MONTHS[m]] || [];
  const qnFactArr = data.daily.qn_fact[MONTHS[m]] || [];
  const today = qnArr[d-1];
  const yest  = d>1 ? qnArr[d-2] : null;
  $('heroQn').textContent = fmt(today);
  if(yest!=null){
    const delta = today - yest;
    const el = $('heroDelta');
    el.textContent = (delta>=0?'▲ +':'▼ ')+fmt(delta);
    el.className = 'delta '+(delta>=0?'up':'down');
  } else {
    $('heroDelta').textContent = '—';
  }
  $('heroFact').textContent = 'Qн факт: '+fmt(qnFactArr[d-1],1);

  // KPI блок
  const sdArr  = data.daily.sdacha[MONTHS[m]] || [];
  const zhArr  = data.daily.dns_in[MONTHS[m]] || [];
  const obvArr = data.daily.obv[MONTHS[m]] || [];
  const lossArr= data.daily.loss[MONTHS[m]] || [];
  const parkArr= data.daily.park[MONTHS[m]] || [];

  $('kSdacha').textContent = fmt(sdArr[d-1]);
  $('kSdachaSub').textContent = 'месяц: '+fmt(sumArr(sdArr));

  $('kZh').textContent = fmt(zhArr[d-1]);
  $('kObv').textContent = 'обв. '+(obvArr[d-1]!=null?fmt(obvArr[d-1],1)+' %':'—');

  $('kLoss').textContent = fmt(lossArr[d-1],1);
  $('kLossSub').textContent = 'месяц: '+fmt(sumArr(lossArr),0);

  $('kPark').textContent = fmt(parkArr[d-1]);
  const prevPark = d>1 ? parkArr[d-2] : null;
  if(prevPark!=null){
    const dlt = parkArr[d-1]-prevPark;
    $('kParkSub').textContent = (dlt>=0?'+':'')+fmt(dlt)+' к вчера';
  } else $('kParkSub').textContent='—';

  // Месячные накопления и прогноз
  $('monthName').textContent = monthName+' '+y;
  const daysInMonth = new Date(y, m+1, 0).getDate();
  $('monthDays').textContent = `${d} из ${daysInMonth} дней`;
  const monthDob = sumArr(qnArr);
  const monthSd  = sumArr(sdArr);
  const avgDob = monthDob/d;
  const forecast = avgDob*daysInMonth;
  $('mDob').textContent = fmt(monthDob);
  $('mSd').textContent  = fmt(monthSd);
  $('mAvg').textContent = fmt(avgDob,0);
  $('mFor').textContent = fmt(forecast,0);

  // График по дням
  drawChart(qnArr.slice(0, d), monthName);
  $('chartRange').textContent = `1–${d} ${MONTH_NAMES_RU[m].toLowerCase()}`;

  // Пар, закачка, газ, ОЗПВ — из месячных
  const mo = data.monthly;
  if(mo && mo.par){
    $('kPar').textContent = fmt(mo.par[m]);
    $('kParSub').textContent = `ППГ-3: ${fmt(mo.pg3[m])} · СПГУ: ${fmt(mo.spgu[m])} · МПГУ: ${fmt(mo.mpgu[m])}`;
    $('kZak').textContent = fmt((mo.zak_int[m]||0)+(mo.zak_ext_sum[m]||0));
    $('kZakSub').textContent = `вн: ${fmtK(mo.zak_int[m])} · сброс: ${fmtK(mo.zak_ext_sum[m])}`;
    $('kGas').textContent = fmt(Math.round((mo.gaz[m]||0)/1000));
    $('kGasSub').textContent = `всего: ${fmtK(mo.gaz[m])} м³`;
    $('kOzpv').textContent = fmt(mo.ozpv[m]);
  }

  // Фонд
  const f = data.fund || {};
  $('fFlow').textContent = fmt(f.flow);
  $('fPump').textContent = fmt(f.pump);
  $('fTotal').textContent = fmt(f.total);
  $('fInFlow').textContent = fmt(f.inFlow);
  $('fInPump').textContent = fmt(f.inPump);
  $('fInTotal').textContent = fmt((f.inFlow||0)+(f.inPump||0));
  $('fInjSt').textContent = fmt(f.injSteamA);
  $('fInjW').textContent  = fmt(f.injWaterA);
  $('fInjT').textContent  = fmt(f.injTotalA);
}

function sumArr(arr){return arr.reduce((s,v)=>s+(typeof v==='number'?v:0),0);}

// === ГРАФИК (canvas, без зависимостей) ===
function drawChart(values, label){
  const canvas = $('chart');
  const dpr = window.devicePixelRatio || 1;
  const w = canvas.clientWidth, h = canvas.clientHeight;
  canvas.width = w*dpr; canvas.height = h*dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr,dpr);
  ctx.clearRect(0,0,w,h);

  const data = values.filter(v=>v!=null);
  if(data.length<2) return;
  const min = Math.min(...data), max = Math.max(...data);
  const pad = (max-min)*0.15 || 100;
  const yMin = min-pad, yMax = max+pad;
  const xPad = 28, yPad = 18;
  const plotW = w-xPad-8, plotH = h-yPad*2;

  // сетка
  ctx.strokeStyle = '#22304d'; ctx.lineWidth=1; ctx.font='10px sans-serif'; ctx.fillStyle='#6e7a99';
  for(let i=0;i<=3;i++){
    const yy = yPad + plotH*i/3;
    ctx.beginPath(); ctx.moveTo(xPad,yy); ctx.lineTo(w-4,yy); ctx.stroke();
    const val = yMax - (yMax-yMin)*i/3;
    ctx.fillText(Math.round(val).toString(), 2, yy+3);
  }

  // линия
  ctx.strokeStyle = '#5b88ff'; ctx.lineWidth=2;
  ctx.beginPath();
  data.forEach((v,i)=>{
    const x = xPad + plotW*i/(data.length-1);
    const y = yPad + plotH*(1-(v-yMin)/(yMax-yMin));
    if(i===0) ctx.moveTo(x,y); else ctx.lineTo(x,y);
  });
  ctx.stroke();

  // заливка под линией
  ctx.lineTo(xPad+plotW, h-yPad);
  ctx.lineTo(xPad, h-yPad);
  ctx.closePath();
  const grad = ctx.createLinearGradient(0,yPad,0,h);
  grad.addColorStop(0, 'rgba(91,136,255,0.35)');
  grad.addColorStop(1, 'rgba(91,136,255,0)');
  ctx.fillStyle = grad; ctx.fill();

  // точки
  ctx.fillStyle = '#5b88ff';
  data.forEach((v,i)=>{
    const x = xPad + plotW*i/(data.length-1);
    const y = yPad + plotH*(1-(v-yMin)/(yMax-yMin));
    ctx.beginPath(); ctx.arc(x,y,2.5,0,Math.PI*2); ctx.fill();
  });

  // подпись последней точки
  const lastV = data[data.length-1];
  const lx = xPad+plotW, ly = yPad+plotH*(1-(lastV-yMin)/(yMax-yMin));
  ctx.fillStyle='#e7eaf3'; ctx.font='bold 11px sans-serif';
  ctx.textAlign='right';
  ctx.fillText(fmt(lastV), lx-4, ly-6);
  ctx.textAlign='left';
}

// === ЗАГРУЗКА / СОХРАНЕНИЕ ===
function save(data){
  try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }catch(e){console.warn('LS err',e);}
}
function load(){
  try{ const s=localStorage.getItem(STORAGE_KEY); return s? JSON.parse(s):null; }catch(e){return null}
}

async function handleFile(file){
  const buf = await file.arrayBuffer();
  try{
    const data = await parseXLSX(buf);
    save(data);
    render(data);
    $('srcInfo').textContent = 'Источник: '+file.name+' · '+new Date().toLocaleString('ru-RU');
    $('modal').hidden = true;
    $('syncBox').hidden = true;
  }catch(e){
    alert('Ошибка чтения файла: '+e.message);
    console.error(e);
  }
}

async function trySyncFromUrl(){
  const url = localStorage.getItem(SYNC_URL_KEY);
  if(!url) return false;
  try{
    const r = await fetch(url, {cache:'no-store'});
    if(!r.ok) throw new Error('HTTP '+r.status);
    const buf = await r.arrayBuffer();
    const data = await parseXLSX(buf);
    save(data);
    render(data);
    $('srcInfo').textContent = 'Источник: авто-синхр. · '+new Date().toLocaleString('ru-RU');
    return true;
  }catch(e){
    console.warn('sync fail',e);
    return false;
  }
}

// === ДЕМО-СНАПШОТ (вшит) ===
const DEMO = {"lastDate":{"y":2026,"m":4,"d":19,"monthName":"Май"},
"daily":{"qn_tov":{"янв":[5601,5604,5614,5624,5631,5634,5636,5641,5648,5653,5660,5662,5662,5656,5652,5653,5654,5655,5656,5657,5658,5659,5660,5661,5662,5663,5664,5665,5666,5667,5668],"фев":[5625,5620,5628,5673,5674,5672,5671,5673,5670,5672,5671,5674,5677,5673,5675,5673,5674,5672,5670,5671,5673,5675,5672,5670,5668,5666,5664,5662,null,null,null],"мар":[5678,5678,5611,5117,4850,5022,4652,5001,5545,5503,5569,5690,5580,5509,5480,5521,5510,5500,5495,5490,5485,5480,5478,5475,5472,5470,5468,5466,5464,5462,5460],"апр":[5720,5602,5487,5485,5500,5535,5582,5584,5579,5582,5412,5580,5578,5580,5396,5578,5580,5582,5584,5586,5588,5590,5592,5594,5596,5598,5600,5602,5604,5606,null],"май":[5659,5500,5581,5554,5500,5570,5601,5522,5520,5597,5698,5458,5400,5536,5520,5508,5502,5469,5550,null,null,null,null,null,null,null,null,null,null,null,null]},
"qn_fact":{"май":[5689.6,5529.7,5611.2,5584,5529.7,5600.1,5631.3,5551.8,5549.8,5627.3,5728.8,5487.5,5429.2,5565.9,5549.8,5567.8,5532,5499,5580,null,null,null,null,null,null,null,null,null,null,null,null]},
"loss":{"май":[30.6,29.7,30.2,30,29.7,30.1,30.3,29.8,29.8,30.3,30.8,29.5,29.2,29.9,29.8,29.8,30,29.5,30,null,null,null,null,null,null,null,null,null,null,null,null]},
"sdacha":{"май":[5631,5650,3500,5400,5550,3200,5500,5600,3300,5520,5650,3100,5400,5550,3250,5500,3400,5450,3821,null,null,null,null,null,null,null,null,null,null,null,null]},
"park":{"май":[35200,35080,37160,37320,37270,39640,39740,39660,41880,41960,41960,44510,44530,44520,46790,46800,49050,49070,38486,null,null,null,null,null,null,null,null,null,null,null,null]},
"dns_in":{"май":[44500,43200,42800,43100,42500,43000,42900,42700,42600,42800,43400,42100,41700,42300,42100,42000,41800,41600,41947,null,null,null,null,null,null,null,null,null,null,null,null]},
"obv":{"май":[85.1,85.3,85.2,85.4,85.5,85.3,85.2,85.3,85.4,85.3,85.1,85.5,85.6,85.4,85.5,85.4,85.4,85.3,85.3,null,null,null,null,null,null,null,null,null,null,null,null]}},
"monthly":{"dob":[175708,159558,169346,166681,105814,0,0,0,0,0,0,0],"sd":[174703,151073,171009,160974,88888,0,0,0,0,0,0,0],"bit":[0,0,45287,104359,42655,0,0,0,0,0,0,0],"zh":[2393388,2159810,2393887,2326093,1473712,0,0,0,0,0,0,0],"park":[12204,12264,19891,17318,22129,0,0,0,0,0,0,0],"pg3":[356369,322587,353035,337815,212571,0,0,0,0,0,0,0],"mpgu":[32620,29416,32667,31448,19992,0,0,0,0,0,0,0],"spgu":[147685,132312,140838,125819,87126,0,0,0,0,0,0,0],"par":[536674,484315,526540,495082,319689,0,0,0,0,0,0,0],"zak_int":[941392,848289,930446.5,902987,556100,0,0,0,0,0,0,0],"zak_ext_sum":[1171090,1084780,1223790,1215500,761620,0,0,0,0,0,0,0],"vlz":[206940,178420,211690,168300,120760,0,0,0,0,0,0,0],"gaz":[38798461,34813750,38151265,35700393,22742293,0,0,0,0,0,0,0],"ozpv":[406060,367538,379452,393300,248770,0,0,0,0,0,0,0]},
"fund":{"flow":24,"pump":3292,"total":3316,"inFlow":16,"inPump":255,"injSteamE":368,"injWaterE":621,"injTotalE":989,"injSteamA":367,"injWaterA":599,"injTotalA":966}};

// Достроим до полной структуры (нули для пустых параметров)
function expandDemo(d){
  for(const k of Object.keys(PARAM_ROWS)){
    if(!d.daily[k]) d.daily[k] = {};
    for(const m of MONTHS) if(!d.daily[k][m]) d.daily[k][m] = Array(31).fill(null);
  }
  return d;
}

// === ОБРАБОТЧИКИ ===
function closeModal(){
  $('modal').hidden = true;
  $('syncBox').hidden = true;
}
function openModal(){
  $('modal').hidden = false;
}

$('fileInput').addEventListener('change', async e=>{
  closeModal();  // <- закрываем меню сразу как пользователь выбрал файл
  const f = e.target.files[0];
  if(f){
    await handleFile(f);
    e.target.value = '';  // сбрасываем, чтобы можно было выбрать тот же файл снова
  }
});
$('refreshBtn').addEventListener('click', async ()=>{
  $('refreshBtn').textContent='…';
  const ok = await trySyncFromUrl();
  $('refreshBtn').textContent='↻';
  if(!ok && !localStorage.getItem(SYNC_URL_KEY)) alert('Ссылка авто-синхронизации не настроена. Меню → Настроить.');
});
$('menuBtn').addEventListener('click', openModal);
$('mClose').addEventListener('click', closeModal);
$('mCloseX').addEventListener('click', closeModal);
$('mUpload').addEventListener('click', ()=>{
  closeModal();
  $('fileInput').click();
});
$('mClear').addEventListener('click', ()=>{
  closeModal();
  if(confirm('Очистить все локальные данные?')){
    localStorage.removeItem(STORAGE_KEY);
    location.reload();
  }
});
$('mSync').addEventListener('click', ()=>{
  $('syncBox').hidden = !$('syncBox').hidden;
  $('syncUrl').value = localStorage.getItem(SYNC_URL_KEY) || '';
});
$('saveSync').addEventListener('click', ()=>{
  const url = $('syncUrl').value.trim();
  if(url) localStorage.setItem(SYNC_URL_KEY, url);
  else localStorage.removeItem(SYNC_URL_KEY);
  closeModal();
  alert('Сохранено. Жми ↻ чтобы обновить.');
});
$('loadDemoBtn').addEventListener('click', ()=>{
  const d = expandDemo(JSON.parse(JSON.stringify(DEMO)));
  save(d); render(d);
  $('srcInfo').textContent='Источник: демо-данные (19.05.2026)';
});
$('modal').addEventListener('click', e=>{ if(e.target.id==='modal') closeModal(); });
// Закрытие по Escape
document.addEventListener('keydown', e=>{ if(e.key==='Escape') closeModal(); });

// === СТАРТ ===
(async()=>{
  const cached = load();
  if(cached){ render(cached); $('srcInfo').textContent='Источник: локальный кэш'; }
  // Попробуем синхр при старте
  const synced = await trySyncFromUrl();
  // Если кэша нет и синхр не сработала — покажем демо
  if(!cached && !synced){
    const d = expandDemo(JSON.parse(JSON.stringify(DEMO)));
    render(d);
    $('srcInfo').textContent='Демо. Загрузите свой XLSX через 📂';
  }
})();
