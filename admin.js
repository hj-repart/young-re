(function(){
const Y = window.YRE, sb = Y.sb, E = Y.esc;
const $ = s => document.querySelector(s);
const main = $('#main');
const NAMES = {board:'상황판',assign:'업무 부여',qr:'QR 만들기',notes:'작업내역',query:'조회',pay:'정산',stmt:'명세서',items:'항목·단가',accounts:'계정',reset:'초기화'};
const TAX = {daily:'일용근로소득',biz33:'3.3% 사업소득',none:'세금 없음'};
const TAXLIST = [{id:'daily',name:'일용근로'},{id:'biz33',name:'3.3%'},{id:'none',name:'없음'}];
let S = {items:[],locs:[],workers:[],prices:[],settings:[]};
let rt = null, st = {};

/* ---------- 공통 ---------- */
async function loadRefs(){
  const r = await Promise.all([
    sb.from('items').select('*').order('sort').order('name'),
    sb.from('locations').select('*').order('sort').order('name'),
    sb.from('workers').select('*').order('name'),
    sb.from('item_prices').select('*').order('effective_from',{ascending:false}),
    sb.from('settings').select('*').order('key'),
    sb.from('worker_prices').select('*').order('effective_from',{ascending:false})]);
  const bad = r.slice(0,5).find(x => x.error);
  if(bad) Y.toast('데이터를 불러오지 못했습니다: ' + bad.error.message, 1);
  S.items = r[0].data||[]; S.locs = r[1].data||[]; S.workers = r[2].data||[]; S.prices = r[3].data||[]; S.settings = r[4].data||[]; S.wprices = r[5].error ? [] : (r[5].data||[]);
}
const activeItems = () => S.items.filter(i => i.active);
const activeLocs = () => S.locs.filter(l => l.active);
const monthEnd = ym => new Date(Date.UTC(+ym.slice(0,4), +ym.slice(5,7), 0)).toISOString().slice(0,10);
const priceOn = (id,d) => { const p = S.prices.find(p => p.item_id===id && String(p.effective_from)<=d); return p ? p.price : 0; };
const priceFor = (wid,id,d) => { const p = (S.wprices||[]).find(p => p.worker_id===wid && p.item_id===id && String(p.effective_from)<=d); return p ? p.price : priceOn(id,d); };
const PAYLIST = [{id:'daily',name:'일급'},{id:'weekly',name:'주급'},{id:'monthly',name:'월급'}], PAYNAME = {daily:'일급',weekly:'주급',monthly:'월급'};
const TAXLIST2 = [{id:'biz33',name:'3.3%'},{id:'none',name:'없음'}];
const setv = k => +((S.settings||[]).find(x => x.key===k)||{}).value || 0;
const mondayOf = d => { const t = new Date(d + 'T00:00:00Z'), w = (t.getUTCDay()+6)%7; t.setUTCDate(t.getUTCDate() - w); return t.toISOString().slice(0,10); };
const opts = (list, sel, blank) => (blank!=null ? `<option value="">${E(blank)}</option>` : '') + list.map(o => `<option value="${o.id}"${o.id===sel?' selected':''}>${E(o.name)}</option>`).join('');
const nameOf = (list,id) => (list.find(x => x.id===id)||{}).name || '';

function cs(id, list, sel, blank){
  const all = (blank!=null ? [{id:'',name:blank}] : []).concat(list);
  const cur = all.some(o => o.id===sel) ? sel : (all[0] ? all[0].id : '');
  return `<div class="cp" id="${id}" data-v="${E(cur)}">${all.map(o => `<button type="button" class="cb${o.id===cur?' on':''}" data-v="${E(o.id)}">${E(o.name)}</button>`).join('')}</div>`;
}
function initChips(root){
  root.querySelectorAll('.cp:not([data-i])').forEach(el => {
    el.dataset.i = 1;
    Object.defineProperty(el, 'value', {get(){ return el.dataset.v; }, set(v){ el.dataset.v = v; el.querySelectorAll('.cb').forEach(b => b.classList.toggle('on', b.dataset.v===String(v))); }});
    el.addEventListener('click', e => { const b = e.target.closest('.cb'); if(!b) return; el.value = b.dataset.v; el.dispatchEvent(new Event('change')); });
  });
}
{ const d = Object.getOwnPropertyDescriptor(Element.prototype,'innerHTML');
  Object.defineProperty(main,'innerHTML',{get(){ return d.get.call(main); }, set(h){ d.set.call(main,h); initChips(main); }}); }
const gcls = g => g==='M' ? 'gm' : g==='F' ? 'gf' : '';
const wn = (name,g) => `<span class="gn ${gcls(g)}">${E(name)}</span>`;
const GLIST = [{id:'M',name:'남'},{id:'F',name:'여'},{id:'',name:'미정'}];
const errMsg = e => /duplicate|unique/i.test(e.message||'') ? '이미 같은 내용이 있습니다' : (e.message||'오류가 발생했습니다');

function modal(title, body, onSave, label){
  const d = document.createElement('dialog'); d.className = 'md';
  d.innerHTML = `<form method="dialog" onsubmit="return false"><h3>${E(title)}</h3><div class="mb">${body}</div><div class="msg" id="mm"></div><div class="mf"><button type="button" class="btn ghost" id="mc">취소</button><button type="button" class="btn" id="ms">${label||'저장'}</button></div></form>`;
  document.body.append(d); initChips(d); d.showModal();
  const close = () => { try{d.close();}catch(e){} d.remove(); };
  d.querySelector('#mc').onclick = close;
  d.addEventListener('cancel', () => setTimeout(() => d.remove()));
  d.querySelector('#ms').onclick = async () => {
    const m = d.querySelector('#mm'); m.textContent = '';
    const r = await onSave(d);
    if(r === true) close(); else if(r) m.textContent = r;
  };
  return d;
}
function xls(name, header, rows){
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '시트1');
  XLSX.writeFile(wb, name + '.xlsx');
}

function nowFor(date){ return date===Y.today() ? new Date().toISOString() : new Date(date+'T09:00:00+09:00').toISOString(); }
async function adminIn(w){
  return sb.from('attendance').insert({worker_id:w.worker_id,work_date:w.work_date,item_id:w.item_id,location_id:w.location_id||null,assignment_id:w.assignment_id||null,check_in:w.check_in||nowFor(w.work_date),price_applied:priceFor(w.worker_id,w.item_id,w.work_date),source:'admin'});
}
async function noteCards(box, itemId, onPick){
  if(!itemId){ box.innerHTML = ''; return; }
  const {data} = await sb.from('work_notes').select('content,location_id,work_date').eq('item_id',itemId).order('work_date',{ascending:false}).limit(60);
  const seen = new Set(), list = [];
  (data||[]).forEach(n => { if(!seen.has(n.content) && list.length<8){ seen.add(n.content); list.push(n); } });
  box.innerHTML = list.length ? list.map((n,i) => `<button type="button" class="nc" data-i="${i}"><b>${E(n.content)}</b><span>${Y.fmtDate(n.work_date)}</span></button>`).join('') : '<div class="muted">이전 작업내역이 없습니다</div>';
  box.querySelectorAll('.nc').forEach(b => b.onclick = () => onPick(list[b.dataset.i]));
}
async function saveNote(date, itemId, locId, text){
  text = (text||'').trim(); if(!text) return;
  const {data} = await sb.from('work_notes').select('id').eq('work_date',date).eq('item_id',itemId).eq('content',text).limit(1);
  if(!data || !data.length) await sb.from('work_notes').insert({work_date:date,item_id:itemId,location_id:locId||null,content:text});
}

/* ---------- 상황판 ---------- */
async function pBoard(){
  const date = st.boardDate || Y.today();
  main.innerHTML = `<h1 class="pt">상황판</h1><div class="cd bdr"><label class="fl"><span>날짜</span><input type="date" id="bd-date" value="${date}"></label><button class="btn line sm" id="bd-today">오늘</button><span class="muted" style="padding-bottom:12px">실시간으로 갱신됩니다</span></div><div id="bd"></div>`;
  const load = async () => {
    const d = $('#bd-date').value || date; st.boardDate = d;
    const [a, g] = await Promise.all([
      sb.from('attendance').select('id,check_in,check_out,price_applied,worker_id,item_id,location_id,workers(name,gender),items(name),locations(name)').eq('work_date',d).order('check_in'),
      sb.from('assignments').select('id,worker_id,item_id,location_id,task_text,workers(name,gender),items(name),locations(name)').eq('work_date',d)]);
    if(a.error || g.error){ $('#bd').innerHTML = `<div class="warn">불러오지 못했습니다: ${E((a.error||g.error).message)}</div>`; return; }
    const att = a.data||[], key = r => r.worker_id + '|' + r.item_id, have = new Set(att.map(key));
    const wait = (g.data||[]).filter(r => !have.has(key(r)));
    const working = att.filter(r => !r.check_out).length, done = att.length - working;
    const cost = att.reduce((s,r) => s + (r.price_applied||0), 0);
    const byItem = new Map(), byLoc = new Map();
    att.forEach(r => { const k = (r.items||{}).name||'-'; (byItem.get(k) || byItem.set(k,[]).get(k)).push({n:(r.workers||{}).name, g:(r.workers||{}).gender, s:r.check_out?'done':'work'}); const l = (r.locations||{}).name||'미지정'; byLoc.set(l,(byLoc.get(l)||0)+1); });
    wait.forEach(r => { const k = (r.items||{}).name||'-'; (byItem.get(k) || byItem.set(k,[]).get(k)).push({n:(r.workers||{}).name, g:(r.workers||{}).gender, s:'wait'}); });
    const max = Math.max(1, ...byLoc.values());
    $('#bd').innerHTML = `<div class="grid g4" style="margin-bottom:14px">
      <div class="met a"><div class="l">출근 인원</div><div class="v">${att.length}명</div></div>
      <div class="met b"><div class="l">근무중 / 퇴근</div><div class="v">${working} / ${done}</div></div>
      <div class="met c"><div class="l">미출근 (업무 부여)</div><div class="v">${wait.length}명</div></div>
      <div class="met"><div class="l">인건비 (총액)</div><div class="v">${Y.won(cost)}</div></div></div>
      <div class="grid g2"><div class="cd"><h3>항목별 인원</h3>${byItem.size ? [...byItem].map(([k,v]) => `<div class="ih"><span>${E(k)}</span><span class="muted">${v.length}명</span></div><div class="chips">${v.map(x => `<span class="pc ${gcls(x.g)} ${x.s}"><i></i>${E(x.n)}</span>`).join('')}</div>`).join('') : '<p class="muted">기록이 없습니다</p>'}
      <p class="muted" style="margin-top:12px">이름 색: <span class="gn gm">남자 파랑</span> · <span class="gn gf">여자 빨강</span> / 점 색: 초록 근무중 · 회색 퇴근 · 점선 미출근</p></div>
      <div class="cd"><h3>장소별 인원</h3>${byLoc.size ? [...byLoc].map(([k,v]) => `<div class="bar"><span class="bl">${E(k)}</span><span class="bt"><i style="width:${Math.round(v/max*100)}%"></i></span><span class="bn">${v}명</span></div>`).join('') : '<p class="muted">기록이 없습니다</p>'}</div></div>
      <div class="cd"><h3>직접 처리 <span class="muted">(QR 없이 관리자가 출퇴근 처리)</span></h3>${(wait.length || working) ? `<div class="prl">${wait.map(r => `<div class="pr"><div class="pn">${wn((r.workers||{}).name,(r.workers||{}).gender)}<span class="muted">${E((r.items||{}).name)}${(r.locations||{}).name?' · '+E(r.locations.name):''}</span></div><span class="chip wait">미출근</span><button class="btn sm" data-in="${r.id}">출근 처리</button></div>`).join('')}${att.filter(r => !r.check_out).map(r => `<div class="pr"><div class="pn">${wn((r.workers||{}).name,(r.workers||{}).gender)}<span class="muted">${E((r.items||{}).name)}${(r.locations||{}).name?' · '+E(r.locations.name):''}</span></div><span class="chip work">근무중 ${Y.fmtTime(r.check_in)}~</span><button class="btn line sm" data-out="${r.id}">퇴근 처리</button></div>`).join('')}</div>` : '<p class="muted">처리할 인원이 없습니다. 업무 부여 메뉴에서 배치하거나 조회 메뉴에서 기록을 추가하세요.</p>'}</div>`;
    const dd = $('#bd-date').value || date;
    $('#bd').querySelectorAll('[data-in]').forEach(b => b.onclick = async () => { const r = (g.data||[]).find(x => x.id===b.dataset.in); b.disabled = true; const x = await adminIn({worker_id:r.worker_id,work_date:dd,item_id:r.item_id,location_id:r.location_id,assignment_id:r.id}); if(x.error){ Y.toast(errMsg(x.error),1); b.disabled = false; } else { Y.toast('출근 처리했습니다'); load(); } });
    $('#bd').querySelectorAll('[data-out]').forEach(b => b.onclick = async () => { b.disabled = true; const x = await sb.from('attendance').update({check_out:nowFor(dd)}).eq('id',b.dataset.out); if(x.error){ Y.toast(errMsg(x.error),1); b.disabled = false; } else { Y.toast('퇴근 처리했습니다'); load(); } });
  };
  $('#bd-date').onchange = load;
  $('#bd-today').onclick = () => { $('#bd-date').value = Y.today(); load(); };
  await load();
  rt = sb.channel('att-board').on('postgres_changes',{event:'*',schema:'public',table:'attendance'}, () => load()).on('postgres_changes',{event:'*',schema:'public',table:'assignments'}, () => load()).subscribe();
}


/* 항목·장소 빠른 추가 (업무 부여 / QR 만들기에서 바로 사용) */
function quickAdd(kind, redraw){
  if(kind==='item'){
    modal('항목 추가', `<label class="fl"><span>항목 이름</span><input id="x-n" maxlength="20" placeholder="예: 야간"></label><label class="fl"><span>일당(원)</span><input id="x-p" type="number" min="0" step="1000" placeholder="0"></label><p class="muted">일당은 나중에 항목·단가 메뉴에서 바꿀 수 있습니다.</p>`, async d => {
      const n = d.querySelector('#x-n').value.trim(), p = parseInt(d.querySelector('#x-p').value,10);
      if(!n) return '항목 이름을 입력하세요';
      const r = await sb.from('items').insert({name:n,sort:S.items.length+1}).select('id').single(); if(r.error) return errMsg(r.error);
      if(p > 0) await sb.from('item_prices').insert({item_id:r.data.id,price:p,effective_from:'2000-01-01'});
      await loadRefs(); Y.toast('항목을 추가했습니다'); redraw(); return true; }, '추가');
  } else {
    modal('장소 추가', `<label class="fl"><span>장소 이름</span><input id="x-n" maxlength="20" placeholder="예: 인천"></label>`, async d => {
      const n = d.querySelector('#x-n').value.trim(); if(!n) return '장소 이름을 입력하세요';
      const r = await sb.from('locations').insert({name:n,sort:S.locs.length+1}); if(r.error) return errMsg(r.error);
      await loadRefs(); Y.toast('장소를 추가했습니다'); redraw(); return true; }, '추가');
  }
}
const qaBtn = k => `<button type="button" class="btn line sm" data-qa="${k}" style="margin-left:8px;min-height:28px;padding:0 10px;font-size:12.5px">+ 추가</button>`;

/* ---------- 업무 부여 ---------- */
async function pAssign(){
  const staff = S.workers;
  main.innerHTML = `<h1 class="pt">업무 부여</h1><div class="cd"><h3>직원에게 업무 부여</h3>
    <div class="row2"><label class="fl"><span>날짜</span><input type="date" id="a-date" value="${st.aDate||Y.today()}"></label>
    <div class="fl"><span>항목${qaBtn('item')}</span>${cs('a-item', activeItems(), st.aItem)}</div>
    <div class="fl"><span>장소${qaBtn('loc')}</span>${cs('a-loc', activeLocs(), st.aLoc, '선택 안 함')}</div></div>
    <div class="ih"><span>배치 기간</span></div>
    <div class="seg" id="a-span"><button type="button" data-s="1" class="on">하루 (선택한 날짜)</button><button type="button" data-s="r">기간 선택</button></div>
    <div id="a-rng" style="display:none;margin-top:8px"><label class="fl"><span>종료일</span><input type="date" id="a-end"></label><label style="display:flex;align-items:center;gap:8px;margin:2px 0 6px;font-size:14px"><input type="checkbox" id="a-wk" checked style="width:18px;height:18px;margin:0"> 토·일 포함</label>
    <p class="muted" id="a-rh" style="margin:4px 0 0">시작일(위 날짜)부터 종료일까지 매일 같은 내용으로 배치합니다. 최대 31일.</p></div>
    <div class="ih"><span>전일 배치</span><span class="acts"><button class="btn line sm" id="a-prev">전일 불러오기</button></span></div>
    <div class="muted" style="margin:-2px 0 6px">전날의 항목·장소·직원·작업 내용을 선택한 날짜로 복사합니다. 이미 같은 항목으로 배치된 직원은 건너뜁니다.</div>
    <div class="ih"><span>근로자 선택</span><span class="acts"><button class="btn line sm" id="a-all">직원 전체</button><button class="btn line sm" id="a-new">알바 추가</button></span></div>
    <div class="muted" style="margin:-2px 0 4px">이름 색: <span class="gn gm">남자 파랑</span> · <span class="gn gf">여자 빨강</span> · 미정은 회색</div><div id="a-staff">${S.workers.length ? [['직원', S.workers.filter(w => w.is_staff)], ['알바', S.workers.filter(w => !w.is_staff)]].filter(g => g[1].length).map(g => `<div class="muted" style="margin:10px 0 6px">${g[0]} ${g[1].length}명</div><div class="wpg">${g[1].map(w => `<label class="wp ${gcls(w.gender)}"><input type="checkbox" value="${w.id}" data-st="${w.is_staff?1:0}"><span>${E(w.name)}</span></label>`).join('')}</div>`).join('') : '<p class="muted">등록된 근로자가 없습니다. 알바 추가로 등록하거나 직원 가입을 승인하세요.</p>'}</div>
    <div class="ih"><span>처리 방식</span></div>
    <div class="seg" id="a-mode"><button type="button" data-m="plan" class="on">배치만 (직원 앱에 표시)</button><button type="button" data-m="in">배치 + 출근 처리</button></div>
    <p class="muted" id="a-mh" style="margin:6px 0 0">직원이 앱에서 직접 출근 버튼을 누릅니다.</p>
    <div class="ih"><span>작업 내용</span></div><div class="ncs" id="a-cards"></div>
    <label class="fl"><textarea id="a-task" placeholder="예: 3층 바닥 청소"></textarea></label>
    <div class="msg" id="a-msg"></div><button class="btn sm" id="a-go" style="margin-top:6px">업무 부여</button></div>
    <div class="cd"><h3>부여 내역 <span class="muted" id="a-day"></span></h3><div id="a-list"></div></div>`;
  const cards = () => noteCards($('#a-cards'), $('#a-item').value, n => { $('#a-task').value = n.content; if(n.location_id && S.locs.find(l=>l.id===n.location_id&&l.active)) $('#a-loc').value = n.location_id; });
  const list = async () => {
    const d = $('#a-date').value; st.aDate = d; $('#a-day').textContent = d ? Y.fmtDate(d) : '';
    const {data,error} = await sb.from('assignments').select('id,task_text,workers(name),items(name),locations(name)').eq('work_date',d).order('created_at');
    if(error){ $('#a-list').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    $('#a-list').innerHTML = (data||[]).length ? `<div class="tw"><table class="t"><thead><tr><th>직원</th><th>항목</th><th>장소</th><th>작업 내용</th><th></th></tr></thead><tbody>${data.map(r => `<tr><td>${E((r.workers||{}).name)}</td><td>${E((r.items||{}).name)}</td><td>${E((r.locations||{}).name||'-')}</td><td style="white-space:normal">${E(r.task_text||'-')}</td><td><button class="btn red sm" data-del="${r.id}">삭제</button></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted">이 날짜에 부여된 업무가 없습니다</p>';
    $('#a-list').querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
      const lk = await sb.from('attendance').select('id').eq('assignment_id',b.dataset.del); if(lk.error){ Y.toast(errMsg(lk.error),1); return; }
      const n = (lk.data||[]).length;
      if(!confirm(n ? '이 업무를 삭제할까요?\n연결된 출퇴근 기록 ' + n + '건도 함께 삭제됩니다.' : '이 업무를 삭제할까요?')) return;
      if(n){ const x = await sb.from('attendance').delete().eq('assignment_id',b.dataset.del); if(x.error){ Y.toast(errMsg(x.error),1); return; } }
      const r = await sb.from('assignments').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast(n ? '업무와 출퇴근 기록을 삭제했습니다' : '삭제했습니다'); list(); } });
  };
  $('#a-prev').onclick = async () => {
    const d = $('#a-date').value; if(!d){ Y.toast('날짜를 선택하세요',1); return; }
    const t = new Date(d + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() - 1); const pd = t.toISOString().slice(0,10);
    const q = await sb.from('assignments').select('worker_id,item_id,location_id,task_text').eq('work_date',pd);
    if(q.error){ Y.toast(errMsg(q.error),1); return; }
    const src = q.data || [];
    if(!src.length){ Y.toast(Y.fmtDate(pd) + ' 배치가 없습니다',1); return; }
    if(!confirm(Y.fmtDate(pd) + ' 배치 ' + src.length + '건을 ' + Y.fmtDate(d) + '로 불러올까요?\n(이미 같은 항목으로 배치된 직원은 건너뜁니다)')) return;
    const ex = await sb.from('assignments').select('worker_id,item_id').eq('work_date',d);
    if(ex.error){ Y.toast(errMsg(ex.error),1); return; }
    const have = new Set((ex.data||[]).map(r => r.worker_id + '|' + r.item_id));
    const add = src.filter(r => !have.has(r.worker_id + '|' + r.item_id)).map(r => ({work_date:d,worker_id:r.worker_id,item_id:r.item_id,location_id:r.location_id,task_text:r.task_text}));
    if(add.length){ const r = await sb.from('assignments').upsert(add,{onConflict:'work_date,worker_id,item_id',ignoreDuplicates:true}); if(r.error){ Y.toast(errMsg(r.error),1); return; } }
    Y.toast(add.length + '건 불러옴' + (src.length - add.length ? ', ' + (src.length - add.length) + '건 건너뜀' : ''));
    list();
  };
  $('#a-item').onchange = () => { st.aItem = $('#a-item').value; cards(); };
  $('#a-date').onchange = list;
  let mode = 'plan';
  let span = '1';
  document.querySelectorAll('#a-span button').forEach(b => b.onclick = () => { span = b.dataset.s; document.querySelectorAll('#a-span button').forEach(x => x.classList.toggle('on', x===b)); $('#a-rng').style.display = span==='r' ? '' : 'none'; if(span==='r' && !$('#a-end').value) $('#a-end').value = $('#a-date').value; });
  document.querySelectorAll('#a-mode button').forEach(b => b.onclick = () => { mode = b.dataset.m; document.querySelectorAll('#a-mode button').forEach(x => x.classList.toggle('on', x===b)); $('#a-mh').textContent = mode==='in' ? '선택한 사람을 관리자가 바로 출근 처리합니다. (오늘이면 지금 시각, 다른 날짜면 09:00 기준)' : '직원이 앱에서 직접 출근 버튼을 누릅니다.'; });
  const syncPick = () => document.querySelectorAll('.wp input').forEach(i => i.closest('.wp').classList.toggle('on', i.checked));
  document.querySelectorAll('.wp input').forEach(i => i.onchange = syncPick);
  const all = $('#a-all'); if(all) all.onclick = () => { document.querySelectorAll('#a-staff input[data-st="1"]').forEach(c => c.checked = true); syncPick(); };
  main.querySelectorAll('[data-qa]').forEach(b => b.onclick = () => quickAdd(b.dataset.qa === 'item' ? 'item' : 'loc', pAssign));
  $('#a-new').onclick = () => modal('알바 추가', `<label class="fl"><span>이름</span><input id="w-name" maxlength="20"></label><label class="fl"><span>전화번호</span><input id="w-phone" type="tel" inputmode="numeric" placeholder="010-0000-0000"></label><div class="fl"><span>성별</span>${cs('w-g', GLIST, '')}</div>`, async d => {
    const n = d.querySelector('#w-name').value.trim(), p = Y.digits(d.querySelector('#w-phone').value);
    if(n.length<2) return '이름을 입력하세요'; if(!Y.validPhone(p)) return '전화번호를 정확히 입력하세요';
    const r = await sb.from('workers').upsert({name:n,phone:p,gender:d.querySelector('#w-g').value||null},{onConflict:'phone',ignoreDuplicates:true}); if(r.error) return errMsg(r.error);
    await loadRefs(); Y.toast('추가했습니다'); pAssign(); return true; });
  $('#a-go').onclick = async () => {
    const m = $('#a-msg'); m.textContent = '';
    const ids = [...document.querySelectorAll('#a-staff input:checked')].map(c => c.value);
    const date = $('#a-date').value, item = $('#a-item').value, loc = $('#a-loc').value || null, task = $('#a-task').value.trim();
    if(!date || !item){ m.textContent = '날짜와 항목을 선택하세요'; return; }
    if(!ids.length){ m.textContent = '직원을 선택하세요'; return; }
    let dates = [date], isR = span === 'r';
    if(isR){
      const end = $('#a-end').value; if(!end){ m.textContent = '종료일을 선택하세요'; return; } if(end < date){ m.textContent = '종료일이 시작일보다 빠릅니다'; return; }
      const wk = $('#a-wk').checked; dates = [];
      for(let t = new Date(date + 'T00:00:00Z'), e = new Date(end + 'T00:00:00Z'); t <= e; t.setUTCDate(t.getUTCDate()+1)){ const g = t.getUTCDay(); if(wk || (g!==0 && g!==6)) dates.push(t.toISOString().slice(0,10)); }
      if(!dates.length){ m.textContent = '배치할 날짜가 없습니다'; return; }
      if(dates.length > 31){ m.textContent = '기간은 최대 31일까지 가능합니다'; return; }
      if(mode==='in' && dates.some(x => x > Y.today())){ m.textContent = '출근 처리는 오늘까지의 날짜에만 할 수 있습니다. 배치만으로 하세요'; return; }
      if(!confirm(Y.fmtDate(dates[0]) + ' ~ ' + Y.fmtDate(dates[dates.length-1]) + ' (' + dates.length + '일) × ' + ids.length + '명을 배치할까요?\n이미 같은 항목으로 배치된 날은 건너뜁니다.')) return;
    }
    $('#a-go').disabled = true;
    const rowsA = dates.flatMap(dt => ids.map(w => ({work_date:dt,worker_id:w,item_id:item,location_id:loc,task_text:task||null})));
    let {data:asg,error} = await sb.from('assignments').upsert(rowsA, isR ? {onConflict:'work_date,worker_id,item_id',ignoreDuplicates:true} : {onConflict:'work_date,worker_id,item_id'}).select('id,worker_id,work_date');
    if(error){ $('#a-go').disabled = false; m.textContent = errMsg(error); return; }
    const added = (asg||[]).length;
    if(isR && mode==='in'){ const q = await sb.from('assignments').select('id,worker_id,work_date').in('work_date',dates).in('worker_id',ids).eq('item_id',item); if(q.error){ $('#a-go').disabled = false; m.textContent = errMsg(q.error); return; } asg = q.data; }
    let done = 0;
    if(mode==='in'){
      const rows = (asg||[]).map(a => ({worker_id:a.worker_id,work_date:a.work_date||date,item_id:item,location_id:loc,assignment_id:a.id,check_in:nowFor(a.work_date||date),price_applied:priceFor(a.worker_id,item,a.work_date||date),source:'admin'}));
      const x = await sb.from('attendance').upsert(rows,{onConflict:'worker_id,work_date,item_id',ignoreDuplicates:true}).select('id');
      if(x.error){ $('#a-go').disabled = false; m.textContent = '배치는 저장했지만 출근 처리에 실패했습니다: ' + errMsg(x.error); list(); return; }
      done = (x.data||[]).length;
    }
    $('#a-go').disabled = false;
    await saveNote(date, item, loc, task);
    st.aItem = item; if(isR){ const tot = dates.length*ids.length; Y.toast(dates.length + '일 × ' + ids.length + '명 중 ' + added + '건 배치' + (tot-added ? ', ' + (tot-added) + '건 건너뜀' : '') + (mode==='in' ? ', ' + done + '건 출근 처리' : '')); } else Y.toast(mode==='in' ? `${ids.length}명 배치, ${done}명 출근 처리했습니다` + (done<ids.length ? ` (${ids.length-done}명은 이미 출근 기록이 있습니다)` : '') : ids.length + '명에게 업무를 부여했습니다'); document.querySelectorAll('#a-staff input').forEach(c => c.checked = false); syncPick(); list(); cards();
  };
  cards(); list();
}

/* ---------- QR 만들기 ---------- */
function qrUrl(token){ return new URL('qr.html?t=' + token, location.href).href; }
async function qrImage(token, lines){
  if(typeof QRCode === 'undefined') throw new Error('QR 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침하세요');
  const c = document.createElement('canvas');
  await QRCode.toCanvas(c, qrUrl(token), {width:300, margin:1, errorCorrectionLevel:'M'});
  const o = document.createElement('canvas'); o.width = 340; o.height = 420;
  const x = o.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0,0,340,420);
  x.drawImage(c, 20, 14, 300, 300);
  x.fillStyle = '#1f2547'; x.textAlign = 'center';
  x.font = '700 22px Pretendard, sans-serif'; x.fillText(lines[0], 170, 348, 310);
  x.font = '600 18px Pretendard, sans-serif'; x.fillText(lines[1], 170, 378, 310);
  x.font = '14px Pretendard, sans-serif'; x.fillStyle = '#7d84a8'; x.fillText(lines[2], 170, 402, 310);
  return o.toDataURL('image/png');
}
async function pQr(){
  main.innerHTML = `<h1 class="pt">QR 만들기</h1><div class="cd"><h3>새 QR 만들기</h3>
    <p class="muted" style="margin:0 0 10px">QR을 찍는 알바는 여기서 설정한 날짜·항목·장소·작업 내용으로 출퇴근이 기록됩니다.</p>
    <div class="row2"><label class="fl"><span>날짜</span><input type="date" id="q-date" value="${st.qDate||Y.today()}"></label>
    <div class="fl"><span>항목${qaBtn('item')}</span>${cs('q-item', activeItems(), st.qItem)}</div>
    <div class="fl"><span>장소${qaBtn('loc')}</span>${cs('q-loc', activeLocs(), st.qLoc, '선택 안 함')}</div></div>
    <div class="ih"><span>작업 내용</span></div><div class="ncs" id="q-cards"></div>
    <label class="fl"><textarea id="q-task" placeholder="예: 신한 현장 정리"></textarea></label>
    <div class="msg" id="q-msg"></div><button class="btn sm" id="q-go" style="margin-top:6px">QR 만들기</button></div>
    <div class="cd"><h3>QR 목록 <span class="muted" id="q-day"></span></h3><label class="fl" style="max-width:200px;margin-bottom:10px"><span>조회 날짜</span><input type="date" id="q-fd" value="${st.qDate||Y.today()}"></label><div class="qrs" id="q-list"></div></div>`;
  const cards = () => noteCards($('#q-cards'), $('#q-item').value, n => { $('#q-task').value = n.content; if(n.location_id && S.locs.find(l=>l.id===n.location_id&&l.active)) $('#q-loc').value = n.location_id; });
  const list = async () => { try{ await list0(); }catch(e){ $('#q-list').innerHTML = `<div class="warn">QR 목록을 불러오지 못했습니다: ${E(e.message||String(e))}</div>`; } };
  const list0 = async () => {
    const d = $('#q-fd').value; $('#q-day').textContent = d ? Y.fmtDate(d) : '';
    const {data,error} = await sb.from('qr_codes').select('id,token,work_date,task_text,active,items(name),locations(name)').eq('work_date',d).order('created_at',{ascending:false});
    if(error){ $('#q-list').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    if(!(data||[]).length){ $('#q-list').innerHTML = '<p class="muted">이 날짜에 만든 QR이 없습니다</p>'; return; }
    let qerr = '';
    const imgs = await Promise.all(data.map(async r => { try{ return await qrImage(r.token, [(r.items||{}).name||'', (r.locations||{}).name||'장소 미지정', Y.fmtDate(r.work_date)]); }catch(e){ qerr = e && e.message ? e.message : String(e); return ''; } }));
    $('#q-list').innerHTML = data.map((r,i) => `<div class="qc${r.active?'':' off'}">${imgs[i] ? `<img src="${imgs[i]}" alt="QR">` : `<div class="warn" style="margin:0 0 8px">QR 그림을 만들지 못했습니다${qerr?': '+E(qerr):''}<br><small>아래 링크 복사로 주소는 쓸 수 있습니다</small></div>`}<div style="margin:8px 0 2px"><b>${E((r.items||{}).name)}</b> · ${E((r.locations||{}).name||'-')} <span class="tag ${r.active?'on':'off'}">${r.active?'사용중':'중지'}</span></div><div class="muted" style="min-height:18px">${E(r.task_text||'')}</div><div class="acts" style="justify-content:center;margin-top:8px"><a class="btn line sm" style="display:inline-flex;align-items:center;text-decoration:none" download="QR_${r.work_date}_${E((r.items||{}).name)}_${E((r.locations||{}).name||'')}.png" href="${imgs[i]}">이미지 저장</a><button class="btn line sm" data-cp="${r.token}">링크 복사</button><button class="btn line sm" data-tg="${r.id}" data-on="${r.active}">${r.active?'사용 중지':'다시 사용'}</button><button class="btn red sm" data-del="${r.id}">삭제</button></div></div>`).join('');
    const box = $('#q-list');
    box.querySelectorAll('[data-cp]').forEach(b => b.onclick = async () => { try{ await navigator.clipboard.writeText(qrUrl(b.dataset.cp)); Y.toast('링크를 복사했습니다'); }catch(e){ prompt('링크를 복사하세요', qrUrl(b.dataset.cp)); } });
    box.querySelectorAll('[data-tg]').forEach(b => b.onclick = async () => { const r = await sb.from('qr_codes').update({active: b.dataset.on!=='true'}).eq('id',b.dataset.tg); if(r.error) Y.toast(errMsg(r.error),1); else list(); });
    box.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('이 QR을 삭제할까요?\n이미 찍힌 출퇴근 기록은 남습니다.')) return; const r = await sb.from('qr_codes').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast('삭제했습니다'); list(); } });
  };
  $('#q-item').onchange = () => { st.qItem = $('#q-item').value; cards(); };
  $('#q-fd').onchange = list;
  main.querySelectorAll('[data-qa]').forEach(b => b.onclick = () => quickAdd(b.dataset.qa === 'item' ? 'item' : 'loc', pQr));
  $('#q-go').onclick = async () => {
    const m = $('#q-msg'); m.textContent = '';
    const date = $('#q-date').value, item = $('#q-item').value, loc = $('#q-loc').value || null, task = $('#q-task').value.trim();
    if(!date || !item){ m.textContent = '날짜와 항목을 선택하세요'; return; }
    $('#q-go').disabled = true;
    const {error} = await sb.from('qr_codes').insert({work_date:date,item_id:item,location_id:loc,task_text:task||null});
    $('#q-go').disabled = false;
    if(error){ m.textContent = errMsg(error); return; }
    await saveNote(date, item, loc, task);
    st.qDate = date; st.qItem = item; $('#q-fd').value = date; Y.toast('QR을 만들었습니다. 아래 QR 목록에서 확인하세요'); await list(); cards(); $('#q-list').scrollIntoView({behavior:'smooth',block:'start'});
  };
  cards(); list();
}

/* ---------- 작업내역 ---------- */
async function pNotes(){
  const to = Y.today(), from = Y.addDays(to,-30);
  main.innerHTML = `<h1 class="pt">작업내역</h1><div class="cd"><h3>항목별 작업내역 작성</h3>
    <div class="row2"><label class="fl"><span>날짜</span><input type="date" id="n-date" value="${to}"></label>
    <div class="fl"><span>항목</span>${cs('n-item', activeItems())}</div>
    <div class="fl"><span>장소</span>${cs('n-loc', activeLocs(), '', '선택 안 함')}</div></div>
    <div class="ih"><span>이전 작업내역 (선택하면 자동 입력)</span></div><div class="ncs" id="n-cards"></div>
    <label class="fl"><textarea id="n-text" placeholder="예: 신한 — 1층 로비 유리 청소"></textarea></label>
    <div class="msg" id="n-msg"></div><button class="btn sm" id="n-go" style="margin-top:6px">저장</button></div>
    <div class="cd"><h3>작성 내역</h3><div class="row2" style="margin-bottom:10px"><label class="fl"><span>시작일</span><input type="date" id="n-from" value="${from}"></label><label class="fl"><span>종료일</span><input type="date" id="n-to" value="${to}"></label></div><div id="n-list"></div></div>`;
  const cards = () => noteCards($('#n-cards'), $('#n-item').value, n => { $('#n-text').value = n.content; if(n.location_id && S.locs.find(l=>l.id===n.location_id&&l.active)) $('#n-loc').value = n.location_id; });
  const list = async () => {
    const {data,error} = await sb.from('work_notes').select('id,work_date,content,items(name),locations(name)').gte('work_date',$('#n-from').value).lte('work_date',$('#n-to').value).order('work_date',{ascending:false}).order('created_at',{ascending:false});
    if(error){ $('#n-list').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    $('#n-list').innerHTML = (data||[]).length ? `<div class="tw"><table class="t"><thead><tr><th>일자</th><th>항목</th><th>장소</th><th>작업 내용</th><th></th></tr></thead><tbody>${data.map(r => `<tr><td>${Y.fmtDate(r.work_date)}</td><td>${E((r.items||{}).name)}</td><td>${E((r.locations||{}).name||'-')}</td><td style="white-space:normal">${E(r.content)}</td><td><button class="btn red sm" data-del="${r.id}">삭제</button></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted">작성된 내역이 없습니다</p>';
    $('#n-list').querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('삭제할까요?')) return; const r = await sb.from('work_notes').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { list(); cards(); } });
  };
  $('#n-item').onchange = cards; $('#n-from').onchange = list; $('#n-to').onchange = list;
  $('#n-go').onclick = async () => {
    const m = $('#n-msg'); m.textContent = '';
    const date = $('#n-date').value, item = $('#n-item').value, loc = $('#n-loc').value || null, text = $('#n-text').value.trim();
    if(!date || !item){ m.textContent = '날짜와 항목을 선택하세요'; return; }
    if(!text){ m.textContent = '작업 내용을 입력하세요'; return; }
    const {error} = await sb.from('work_notes').insert({work_date:date,item_id:item,location_id:loc,content:text});
    if(error){ m.textContent = errMsg(error); return; }
    $('#n-text').value = ''; Y.toast('저장했습니다'); list(); cards();
  };
  cards(); list();
}

/* ---------- 조회 ---------- */
function attBody(a){
  return `<div class="fl"><span>항목</span>${cs('e-item', S.items, a.item_id)}</div>
  <div class="fl"><span>장소</span>${cs('e-loc', S.locs, a.location_id, '선택 안 함')}</div>
  <label class="fl"><span>출근</span><input type="datetime-local" id="e-in" value="${Y.toLocalInput(a.check_in)}"></label>
  <label class="fl"><span>퇴근</span><input type="datetime-local" id="e-out" value="${Y.toLocalInput(a.check_out)}"></label>
  <label class="fl"><span>일당(원)</span><input type="number" id="e-price" min="0" step="1000" value="${a.price_applied||0}"></label>
  <label class="fl"><span>메모</span><input type="text" id="e-note" maxlength="100" value="${E(a.note||'')}"></label>`;
}
async function pQuery(){
  const to = Y.today(), q = st.q = st.q || {from: to.slice(0,8)+'01', to, worker:'', item:'', loc:'', view:'person'};
  main.innerHTML = `<h1 class="pt">조회</h1><div class="cd"><div class="row2">
    <label class="fl"><span>시작일</span><input type="date" id="f-from" value="${q.from}"></label><label class="fl"><span>종료일</span><input type="date" id="f-to" value="${q.to}"></label>
    <div class="fl"><span>사람 <small class="muted">(이름 검색 또는 선택)</small></span><input type="search" id="f-name" placeholder="이름 일부를 입력 (예: 김)" value="${E(q.name||'')}" style="height:42px;border-radius:12px;border:1px solid #e1e6fb;background:#f4f6ff;padding:0 12px;font:inherit;font-size:15px;margin-bottom:8px">${cs('f-worker', S.workers, q.worker, '전체')}</div>
    <div class="fl"><span>항목</span>${cs('f-item', S.items, q.item, '전체')}</div>
    <div class="fl"><span>장소</span>${cs('f-loc', S.locs, q.loc, '전체')}</div></div>
    <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:12px"><div class="seg" id="f-view"><button data-v="person">사람별</button><button data-v="date">일자별</button><button data-v="item">항목별</button></div><span class="sp" style="flex:1"></span><button class="btn line sm" id="f-add">기록 추가</button><button class="btn sm" id="f-xls">엑셀 저장</button></div></div>
    <div class="cd"><div id="f-res"></div></div>`;
  let rows = [];
  const draw = () => {
    document.querySelectorAll('#f-view button').forEach(b => b.classList.toggle('on', b.dataset.v===q.view));
    if(!rows.length){ $('#f-res').innerHTML = '<p class="muted">조건에 맞는 기록이 없습니다</p>'; return; }
    const keyOf = r => q.view==='person' ? ((r.workers||{}).name||'-') : q.view==='date' ? Y.fmtDate(r.work_date) : ((r.items||{}).name||'-');
    const groups = new Map(); rows.forEach(r => { const k = keyOf(r); (groups.get(k) || groups.set(k,[]).get(k)).push(r); });
    const keys = [...groups.keys()].sort((a,b) => q.view==='date' ? (a<b?-1:1) : a.localeCompare(b,'ko'));
    const total = rows.reduce((s,r) => s + (r.price_applied||0), 0);
    const itc = id => { const i = S.items.findIndex(x => x.id===id); return 'c' + (i<0?0:i%8); };
    const line = r => { const w = r.workers||{}, main1 = q.view==='date' ? wn(w.name,w.gender) : Y.fmtDate(r.work_date), sub = q.view==='person' ? '' : '';
      return `<div class="qr"><div class="qm"><b>${main1}</b>${q.view==='item' ? `<span class="qs">${wn(w.name,w.gender)}</span>` : ''}</div><div class="qcc"><span class="ti ${itc(r.item_id)}">${E((r.items||{}).name)}</span>${(r.locations||{}).name?`<span class="tl">${E(r.locations.name)}</span>`:''}</div><details class="qt"><summary>${r.check_out ? Y.fmtDur(r.check_in,r.check_out) : '<span class="chip work">근무중</span>'}</summary><span class="qi">출근 ${Y.fmtTime(r.check_in)} · 퇴근 ${r.check_out?Y.fmtTime(r.check_out):'-'}</span></details><div class="qp">${Y.won(r.price_applied)}</div><div class="acts"><button class="btn line sm" data-ed="${r.id}">수정</button><button class="btn red sm" data-del="${r.id}">삭제</button></div></div>`; };
    $('#f-res').innerHTML = `<div class="qtot"><span>전체 ${rows.length}건</span><b>${Y.won(total)}</b></div><div style="display:flex;justify-content:flex-end;margin:0 0 10px"><button class="btn red sm" id="f-delall">조회된 ${rows.length}건 전체 삭제</button></div>` + keys.map(k => { const g = groups.get(k), sum = g.reduce((s,r) => s + (r.price_applied||0), 0), w0 = (g[0].workers||{}); const ttl = q.view==='person' ? wn(k,w0.gender) : E(k); return `<div class="qg"><div class="qh"><span class="qn">${ttl}</span><span class="qk">${g.length}건</span><b>${Y.won(sum)}</b></div>${g.map(line).join('')}</div>`; }).join('');
    $('#f-delall').onclick = async () => {
      const cnt = rows.length, per = q.from === q.to ? Y.fmtDate(q.from) : Y.fmtDate(q.from) + ' ~ ' + Y.fmtDate(q.to);
      if(!confirm('조회된 출퇴근 기록 ' + cnt + '건(' + per + ')을 모두 삭제할까요?\n정산 금액에서도 빠지며 되돌릴 수 없습니다.')) return;
      if(!confirm('정말 ' + cnt + '건을 삭제합니다. 계속할까요?')) return;
      const ids = rows.map(r => r.id); let fail = null;
      for(let i = 0; i < ids.length && !fail; i += 100){ const r = await sb.from('attendance').delete().in('id', ids.slice(i, i+100)); if(r.error) fail = r.error; }
      if(fail) Y.toast(errMsg(fail),1); else Y.toast(cnt + '건을 삭제했습니다');
      load();
    };
    $('#f-res').querySelectorAll('[data-ed]').forEach(b => b.onclick = () => edit(rows.find(r => r.id===b.dataset.ed)));
    $('#f-res').querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('이 출퇴근 기록을 삭제할까요?')) return; const r = await sb.from('attendance').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast('삭제했습니다'); load(); } });
  };
  const load = async () => {
    q.from = $('#f-from').value; q.to = $('#f-to').value; q.worker = $('#f-worker').value; q.item = $('#f-item').value; q.loc = $('#f-loc').value;
    let b = sb.from('attendance').select('id,work_date,check_in,check_out,price_applied,note,worker_id,item_id,location_id,workers(name,phone,gender),items(name),locations(name)').gte('work_date',q.from).lte('work_date',q.to).order('work_date').order('check_in').limit(5000);
    if(q.worker) b = b.eq('worker_id',q.worker); if(q.item) b = b.eq('item_id',q.item); if(q.loc) b = b.eq('location_id',q.loc);
    const {data,error} = await b;
    if(error){ $('#f-res').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    q.name = $('#f-name').value.trim(); rows = (data||[]).filter(r => !q.name || ((r.workers||{}).name||'').includes(q.name)); draw();
  };
  const edit = a => modal('출퇴근 기록 수정', `<label class="fl"><span>근무일</span><input type="date" id="e-date" value="${String(a.work_date).slice(0,10)}"></label>` + attBody(a), async d => {
    const g = s => d.querySelector(s).value;
    const upd = {work_date:g('#e-date'), item_id:g('#e-item'), location_id:g('#e-loc')||null, check_in:Y.fromLocalInput(g('#e-in')), check_out:Y.fromLocalInput(g('#e-out')), price_applied:parseInt(g('#e-price'),10)||0, note:g('#e-note')||null};
    if(!upd.work_date) return '근무일을 입력하세요';
    if(!upd.check_in) return '출근 시간을 입력하세요';
    if(upd.check_out && upd.check_out < upd.check_in) return '퇴근 시간이 출근보다 빠릅니다';
    const r = await sb.from('attendance').update(upd).eq('id',a.id);
    if(r.error) return errMsg(r.error); Y.toast('수정했습니다'); load(); return true;
  });
  $('#f-add').onclick = () => {
    const d0 = q.to || to;
    modal('출퇴근 기록 추가', `<div class="fl"><span>근로자</span>${cs('n-worker', S.workers,'','선택')}</div><label class="fl"><span>근무일</span><input type="date" id="n-d" value="${d0}"></label>` + attBody({item_id:'',location_id:'',check_in:null,check_out:null,price_applied:0}), async d => {
      const g = s => d.querySelector(s).value;
      const w = g('#n-worker'), day = g('#n-d'), item = g('#e-item'), cin = Y.fromLocalInput(g('#e-in')), cout = Y.fromLocalInput(g('#e-out'));
      if(!w || !day || !item) return '근로자, 근무일, 항목을 선택하세요';
      if(!cin) return '출근 시간을 입력하세요';
      if(cout && cout < cin) return '퇴근 시간이 출근보다 빠릅니다';
      const price = g('#e-price') !== '' && parseInt(g('#e-price'),10) > 0 ? parseInt(g('#e-price'),10) : priceFor(w, item, day);
      const r = await sb.from('attendance').insert({worker_id:w,work_date:day,item_id:item,location_id:g('#e-loc')||null,check_in:cin,check_out:cout,price_applied:price,source:'admin',note:g('#e-note')||null});
      if(r.error) return errMsg(r.error); Y.toast('추가했습니다'); load(); return true;
    }, '추가');
  };
  document.querySelectorAll('#f-view button').forEach(b => b.onclick = () => { q.view = b.dataset.v; draw(); });
  ['f-from','f-to','f-worker','f-item','f-loc'].forEach(id => $('#'+id).onchange = load);
  $('#f-name').oninput = () => { clearTimeout(window._qn); window._qn = setTimeout(load, 250); };
  $('#f-xls').onclick = () => { if(!rows.length){ Y.toast('저장할 기록이 없습니다',1); return; } xls(`조회_${q.from}_${q.to}`, ['일자','이름','전화번호','항목','장소','출근','퇴근','근무시간','일당'], rows.map(r => [String(r.work_date).slice(0,10),(r.workers||{}).name,Y.fmtPhone((r.workers||{}).phone),(r.items||{}).name,(r.locations||{}).name||'',Y.fmtTime(r.check_in),r.check_out?Y.fmtTime(r.check_out):'',Y.fmtDur(r.check_in,r.check_out),r.price_applied||0])); };
  load();
}

/* ---------- 정산 ---------- */
async function pPay(){
  const ym = st.payYm || Y.today().slice(0,7);
  main.innerHTML = `<h1 class="pt">정산</h1><div class="cd bdr"><label class="fl"><span>정산 월</span><input type="month" id="p-ym" value="${ym}"></label><span style="flex:1"></span><button class="btn sm" id="p-xls">엑셀 저장</button></div><div id="p-warn"></div><div class="cd"><div id="p-res"></div></div>
    <div class="cd"><h3>계산 기준</h3><p class="muted" style="margin:0;line-height:1.7">일용근로소득: (일당 − 공제액) × 세율 × (1 − 세액공제율)을 하루마다 계산하고, 한 달 합계 세액이 소액부징수 기준 미만이면 0원입니다. 지방소득세는 소득세의 10%(10원 미만 절사)입니다. 3.3% 사업소득은 소득세 3% + 지방소득세 0.3%입니다. 세율은 항목·단가 메뉴에서 바꿀 수 있고, 세금 방식은 계정 메뉴에서 근로자별로 고릅니다.</p></div>`;
  let rows = [];
  const load = async () => {
    const v = $('#p-ym').value; st.payYm = v;
    const [r, z] = await Promise.all([
      stmtData(v).then(d => ({data:d.list.map(x => x.pay)}), e => ({error:e})),
      sb.from('attendance').select('id',{count:'exact',head:true}).eq('price_applied',0).gte('work_date',v+'-01').lte('work_date',monthEnd(v))]);
    if(r.error){ $('#p-res').innerHTML = `<div class="warn">${E(r.error.message)}</div>`; return; }
    rows = r.data||[];
    $('#p-warn').innerHTML = z.count ? `<div class="warn">이 달에 일당이 0원인 기록이 ${z.count}건 있습니다. 항목·단가 메뉴에서 단가를 입력한 뒤 조회 메뉴에서 해당 기록의 일당을 수정하세요.</div>` : '';
    if(!rows.length){ $('#p-res').innerHTML = '<p class="muted">이 달의 기록이 없습니다</p>'; return; }
    const sum = k => rows.reduce((s,x) => s + (x[k]||0), 0);
    $('#p-res').innerHTML = `<div class="tw"><table class="t"><thead><tr><th>이름</th><th>세금 방식</th><th class="r">근무일수</th><th class="r">총 지급액</th><th class="r">소득세</th><th class="r">지방소득세</th><th class="r">고용보험</th><th class="r">실지급액</th></tr></thead><tbody>${rows.map(x => `<tr><td>${wn(x.worker_name,(S.workers.find(w => w.id===x.worker_id)||{}).gender)}</td><td>${TAX[x.tax_type]||x.tax_type}</td><td class="r">${x.days}일</td><td class="r">${Y.won(x.gross)}</td><td class="r">${Y.won(x.income_tax)}</td><td class="r">${Y.won(x.local_tax)}</td><td class="r">${Y.won(x.emp_ins)}</td><td class="r"><b>${Y.won(x.net_pay)}</b></td></tr>`).join('')}<tr class="tot"><td colspan="2">합계 · ${rows.length}명</td><td class="r">${sum('days')}일</td><td class="r">${Y.won(sum('gross'))}</td><td class="r">${Y.won(sum('income_tax'))}</td><td class="r">${Y.won(sum('local_tax'))}</td><td class="r">${Y.won(sum('emp_ins'))}</td><td class="r">${Y.won(sum('net_pay'))}</td></tr></tbody></table></div>`;
  };
  $('#p-ym').onchange = load;
  $('#p-xls').onclick = () => { if(!rows.length){ Y.toast('저장할 내용이 없습니다',1); return; } xls('정산_'+st.payYm, ['이름','전화번호','세금 방식','근무일수','총 지급액','소득세','지방소득세','고용보험','실지급액'], rows.map(x => [x.worker_name,Y.fmtPhone(x.phone),TAX[x.tax_type]||x.tax_type,x.days,x.gross,x.income_tax,x.local_tax,x.emp_ins,x.net_pay])); };
  load();
}

/* ---------- 명세서 ---------- */
const SOPT_KEY = 'yre_stmt_opts', SDEF = {mode:'detail', loc:true, time:false, tax:true, adj:true};
const sopt = () => { try{ return Object.assign({}, SDEF, JSON.parse(localStorage.getItem(SOPT_KEY)||'{}')); }catch(e){ return Object.assign({}, SDEF); } };
const ssave = o => { try{ localStorage.setItem(SOPT_KEY, JSON.stringify(o)); }catch(e){} };
const TAXNAME = {daily:'일용근로소득', biz33:'3.3% 사업소득', none:'세금 없음'};
const ymLabel = ym => ym.slice(0,4) + '년 ' + parseInt(ym.slice(5),10) + '월';
const money = n => (n<0 ? '-' : '') + Y.won(Math.abs(n));

async function stmtData(ym){
  const [m, a, j] = await Promise.all([
    sb.from('v_month_pay').select('*').eq('ym',ym).order('worker_name'),
    sb.from('v_attendance_pay').select('worker_id,work_date,item_name,location_name,check_in,check_out,gross').gte('work_date',ym+'-01').lte('work_date',monthEnd(ym)).order('work_date'),
    sb.from('pay_adjustments').select('*').eq('ym',ym).order('created_at')]);
  if(m.error) throw m.error; if(a.error) throw a.error;
  const byW = new Map();
  (m.data||[]).forEach(r => byW.set(r.worker_id, {pay:r, days:[], adj:[], type:'daily'}));
  (a.data||[]).forEach(r => { const x = byW.get(r.worker_id); if(x) x.days.push(r); });
  S.workers.filter(w => w.pay_type==='weekly' || w.pay_type==='monthly').forEach(w => {
    const days = (a.data||[]).filter(r => r.worker_id===w.id), wk = [...new Set(days.map(r => mondayOf(r.work_date)))].sort();
    const wage = w.pay_type==='weekly' ? (w.weekly_wage||0) : (w.monthly_wage||0);
    if(w.pay_type==='weekly' && !wk.length){ byW.delete(w.id); return; }
    const lines = w.pay_type==='weekly' ? wk.map((d,i) => [`${i+1}주차 주급`, wage]) : [['월급', wage]];
    const gross = lines.reduce((t,l) => t + l[1], 0), tt = w.tax_type==='biz33' ? 'biz33' : 'none';
    const it = tt==='biz33' ? Math.floor(gross * setv('biz_rate')) : 0, lt = Math.floor(it * setv('local_rate') / 10) * 10;
    byW.set(w.id, {type:w.pay_type, lines, days, adj:[], pay:{worker_id:w.id, worker_name:w.name, phone:w.phone, tax_type:tt, days:days.length, weeks:wk.length, gross, income_tax:it, local_tax:lt, emp_ins:0, net_pay:gross - it - lt}});
  });
  (j.data||[]).forEach(r => { const x = byW.get(r.worker_id); if(x) x.adj.push(r); });
  const list = [...byW.values()].sort((p,q) => (p.pay.worker_name||'').localeCompare(q.pay.worker_name||'', 'ko'));
  return {list, adjMissing: !!j.error};
}
function stmtModel(x, ym, o){
  const p = x.pay, yy = +ym.slice(0,4), mm = +ym.slice(5,7), last = new Date(Date.UTC(yy,mm,0)).getUTCDate(), firstDow = new Date(Date.UTC(yy,mm-1,1)).getUTCDay();
  const byItem = new Map(); x.days.forEach(d => { const e = byItem.get(d.item_name) || {n:0,amt:0}; e.n++; e.amt += d.gross; byItem.set(d.item_name, e); });
  const pay = x.type==='daily' ? [...byItem].map(([k,e]) => [k + ' × ' + e.n + '일', e.amt]) : x.lines.map(l => l.slice()), ded = [];
  if(o.adj) x.adj.forEach(r => (r.amount >= 0 ? pay : ded).push([r.label, Math.abs(r.amount)]));
  const taxOn = o.tax && p.tax_type !== 'none';
  if(taxOn){ ded.push(['소득세', p.income_tax||0], ['지방소득세', p.local_tax||0]); if(p.emp_ins) ded.push(['고용보험', p.emp_ins]); }
  const sum = l => l.reduce((t,r) => t + r[1], 0), payT = sum(pay), dedT = sum(ded), sub = x.type==='weekly' ? `주급 · ${p.weeks}주` : x.type==='monthly' ? '월급' : `근무 ${p.days}일`;
  const mark = {}; x.days.forEach(d => { const n = +d.work_date.slice(8,10); mark[n] = (mark[n]||0) + 1; });
  const calc = p.tax_type==='daily' ? '일용근로소득세 = (일당 − 공제액) × 세율 × (1 − 세액공제율), 월 합계 세액이 소액부징수 기준 미만이면 0원. 지방소득세는 소득세의 10%.' : p.tax_type==='biz33' ? '사업소득 3.3% (소득세 3% + 지방소득세 0.3%)' : '세금 공제 없음';
  return {title:'급여명세서', ym:ymLabel(ym), range:`${ym}-01 ~ ${ym}-${String(last).padStart(2,'0')}`, name:p.worker_name, phone:p.phone, taxType:TAXNAME[p.tax_type]||'', days:p.days, sub, payType:PAYNAME[x.type],
    pay, ded, payT, dedT, net:payT - dedT, netLabel: o.tax ? '차인지급액' : '차인지급액(세전)', mark, last, firstDow, showCal:o.mode==='detail' && (x.type==='daily' || x.days.length>0), calc: o.tax ? calc : '',
    showDetail:o.mode==='detail' && x.type==='daily', showTime:!!o.time, rows: o.mode==='detail' && x.type==='daily' ? x.days.map(d => ({date:Y.fmtDate(d.work_date), what:d.item_name + (o.loc && d.location_name ? ' · ' + d.location_name : ''), time:o.time ? Y.fmtDur(d.check_in,d.check_out) : '', amt:d.gross})) : []};
}
const WDN = ['일','월','화','수','목','금','토'];
function stmtHtml(m){
  const n = Math.max(m.pay.length, m.ded.length, 2), pr = i => m.pay[i] ? `<td>${E(m.pay[i][0])}</td><td class="r">${Y.won(m.pay[i][1])}</td>` : '<td></td><td></td>', dr = i => m.ded[i] ? `<td>${E(m.ded[i][0])}</td><td class="r neg">${Y.won(m.ded[i][1])}</td>` : '<td></td><td></td>';
  const cal = []; for(let i=0;i<m.firstDow;i++) cal.push('<i></i>'); for(let d=1; d<=m.last; d++){ const w = (m.firstDow + d - 1) % 7; cal.push(`<i class="${w===0?'su':w===6?'sa':''}${m.mark[d]?' on':''}"><b>${d}</b><span>${m.mark[d]||''}</span></i>`); }
  return `<div class="sheet"><div class="sh-top"><div class="sh-t">${E(m.title)}</div><div class="sh-ym"><small>귀속연월</small>${E(m.ym)}</div></div>
  <div class="sh-info"><div class="sh-who"><div class="sh-nm">${E(m.name)}</div><div class="sh-sub">${E(m.taxType)} · ${E(m.sub)}</div><div class="sh-sub">${E(m.range)}</div></div><div class="sh-net"><small>${E(m.netLabel)}</small><b>${Y.won(m.net)}</b><span>지급합계 ${Y.won(m.payT)}</span><span>공제합계 ${Y.won(m.dedT)}</span></div></div>
  <table class="sh-pd"><thead><tr><th colspan="2">지급내역</th><th colspan="2">공제내역</th></tr></thead><tbody>${Array.from({length:n},(_,i) => `<tr>${pr(i)}${dr(i)}</tr>`).join('')}<tr class="sum"><td>지급합계</td><td class="r">${Y.won(m.payT)}</td><td>공제합계</td><td class="r neg">${Y.won(m.dedT)}</td></tr></tbody></table>
  <div class="sh-bar"><span>${E(m.netLabel)}</span><b>${Y.won(m.net)}</b></div>
  ${m.showCal ? `<div class="sh-h">월간 근무현황</div><div class="sh-cal">${WDN.map((w,i) => `<em class="${i===0?'su':i===6?'sa':''}">${w}</em>`).join('')}${cal.join('')}</div>` : ''}
  ${m.showDetail && m.rows.length ? `<div class="sh-h">근무 내역</div><table class="sh-d"><tbody>${m.rows.map(r => `<tr><td>${E(r.date)}</td><td>${E(r.what)}</td>${m.showTime?`<td>${E(r.time)}</td>`:''}<td class="r">${Y.won(r.amt)}</td></tr>`).join('')}</tbody></table>` : ''}
  ${m.calc ? `<div class="sh-note">${E(m.calc)}</div>` : ''}</div>`;
}
function stmtCanvas(m){
  const W = 800, P = 36, RH = 38, n = Math.max(m.pay.length, m.ded.length, 2), cw = (W - 2*P) / 7, crow = Math.ceil((m.firstDow + m.last) / 7);
  const H = 120 + 130 + 44 + n*RH + RH + 24 + 70 + (m.showCal ? 50 + 30 + crow*52 : 0) + (m.showDetail && m.rows.length ? 50 + m.rows.length*34 : 0) + (m.calc ? 70 : 0) + 30;
  const c = document.createElement('canvas'), K = 2; c.width = W*K; c.height = H*K; const x = c.getContext('2d'); x.scale(K,K);
  const F = (w,sz) => `${w} ${sz}px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif`, INK = '#1f2547', LINE = '#c9d2f5';
  const T = (t, px, py, font, col, al) => { x.font = font; x.fillStyle = col; x.textAlign = al || 'left'; x.fillText(t, px, py); };
  const L = (x1,y1,x2,y2,w,col) => { x.strokeStyle = col || LINE; x.lineWidth = w || 1; x.beginPath(); x.moveTo(x1,y1); x.lineTo(x2,y2); x.stroke(); };
  x.fillStyle = '#fff'; x.fillRect(0,0,W,H);
  let y = 58; T(m.title, P, y, F(800,32), INK); T('귀속연월', W-P, y-18, F(500,13), '#7d84a8', 'right'); T(m.ym, W-P, y+2, F(700,18), INK, 'right'); y += 16; L(P,y,W-P,y,3,INK);
  y += 22; x.fillStyle = '#f4f6ff'; x.fillRect(P, y, W-2*P, 104);
  T(m.name, P+18, y+38, F(800,26), INK); T(`${m.taxType} · ${m.sub}`, P+18, y+64, F(500,15), '#5a6085'); T(m.range, P+18, y+86, F(500,14), '#7d84a8');
  T(m.netLabel, W-P-18, y+28, F(600,14), '#7d84a8', 'right'); T(Y.won(m.net), W-P-18, y+62, F(800,30), '#2c3470', 'right'); T(`지급 ${Y.won(m.payT)}  ·  공제 ${Y.won(m.dedT)}`, W-P-18, y+88, F(500,13), '#7d84a8', 'right');
  y += 104 + 28; const mid = W/2;
  x.fillStyle = '#eef2ff'; x.fillRect(P, y, W-2*P, 36); T('지급내역', P+14, y+24, F(700,16), INK); T('공제내역', mid+14, y+24, F(700,16), INK); y += 36;
  for(let i=0;i<n;i++){ L(P,y+RH,W-P,y+RH,1,'#e3e7fa'); if(m.pay[i]){ T(m.pay[i][0], P+14, y+25, F(500,16), '#3a4066'); T(Y.won(m.pay[i][1]), mid-14, y+25, F(600,16), INK, 'right'); } if(m.ded[i]){ T(m.ded[i][0], mid+14, y+25, F(500,16), '#3a4066'); T(Y.won(m.ded[i][1]), W-P-14, y+25, F(600,16), '#c0275a', 'right'); } y += RH; }
  L(P,y,W-P,y,1.5,LINE); x.fillStyle = '#f8f9ff'; x.fillRect(P, y, W-2*P, RH); T('지급합계', P+14, y+25, F(700,16), INK); T(Y.won(m.payT), mid-14, y+25, F(700,16), INK, 'right'); T('공제합계', mid+14, y+25, F(700,16), INK); T(Y.won(m.dedT), W-P-14, y+25, F(700,16), '#c0275a', 'right'); L(mid, y-n*RH-36, mid, y+RH, 1, LINE); y += RH;
  L(P,y,W-P,y,1.5,LINE); y += 24; x.fillStyle = '#e4eaff'; x.fillRect(P, y, W-2*P, 56); T(m.netLabel, P+18, y+36, F(800,19), '#2c3470'); T(Y.won(m.net), W-P-18, y+38, F(800,28), '#2c3470', 'right'); y += 56;
  if(m.showCal){ y += 38; T('월간 근무현황', P, y, F(700,16), INK); y += 14; WDN.forEach((w,i) => T(w, P + cw*i + cw/2, y+18, F(600,13), i===0?'#c0275a':i===6?'#3b57c8':'#7d84a8', 'center')); y += 26;
    for(let d=1; d<=m.last; d++){ const idx = m.firstDow + d - 1, col = idx % 7, row = Math.floor(idx / 7), cx = P + cw*col, cy = y + row*52; x.strokeStyle = '#e3e7fa'; x.lineWidth = 1; x.strokeRect(cx+1, cy+1, cw-2, 50);
      if(m.mark[d]){ x.fillStyle = '#e8efff'; x.fillRect(cx+2, cy+2, cw-4, 48); } T(String(d), cx+8, cy+18, F(500,12), col===0?'#c0275a':col===6?'#3b57c8':'#7d84a8'); if(m.mark[d]) T(String(m.mark[d]), cx+cw/2, cy+40, F(800,18), '#2f4fc4', 'center'); }
    y += crow*52; }
  if(m.showDetail && m.rows.length){ y += 38; T('근무 내역', P, y, F(700,16), INK); y += 10; m.rows.forEach(r => { y += 34; T(r.date, P+6, y-8, F(500,15), '#4a5068'); T(r.what, P+170, y-8, F(500,15), INK); if(m.showTime) T(r.time, P+450, y-8, F(500,15), '#7d84a8'); T(Y.won(r.amt), W-P-6, y-8, F(600,15), INK, 'right'); L(P,y,W-P,y,1,'#eef0fb'); }); }
  if(m.calc){ y += 30; x.font = F(500,12); x.fillStyle = '#8b90ad'; x.textAlign = 'left'; const words = m.calc.split(' '); let ln = ''; words.forEach(wd => { if(x.measureText(ln + wd).width > W - 2*P){ T(ln, P, y, F(500,12), '#8b90ad'); y += 18; ln = ''; } ln += wd + ' '; }); T(ln, P, y, F(500,12), '#8b90ad'); }
  return c;
}
function stmtSaveImg(m){ const c = stmtCanvas(m), a = document.createElement('a'); a.href = c.toDataURL('image/png'); a.download = `${m.name}_${m.ym}_급여명세서.png`; document.body.append(a); a.click(); a.remove(); }
function stmtSheetRows(m){
  const n = Math.max(m.pay.length, m.ded.length), out = [[m.title, '', '귀속연월', m.ym], [m.name, m.taxType, '급여 유형', m.payType, '근무일수', m.days + '일'], ['지급기간', m.range], [], ['지급내역','금액','공제내역','금액']];
  for(let i=0;i<n;i++) out.push([m.pay[i]?m.pay[i][0]:'', m.pay[i]?m.pay[i][1]:'', m.ded[i]?m.ded[i][0]:'', m.ded[i]?m.ded[i][1]:'']);
  out.push(['지급합계', m.payT, '공제합계', m.dedT], [m.netLabel, m.net], []);
  if(m.showDetail && m.rows.length){ out.push(m.showTime ? ['일자','항목','근무시간','일당'] : ['일자','항목','일당']); m.rows.forEach(r => out.push(m.showTime ? [r.date,r.what,r.time,r.amt] : [r.date,r.what,r.amt])); }
  return out;
}
async function pStmt(){
  const ym = st.stYm || Y.today().slice(0,7); let o = sopt(), cur = null, data = null;
  main.innerHTML = `<h1 class="pt">명세서</h1><div id="st-box"></div>`;
  const draw = async () => {
    const box = $('#st-box'); box.innerHTML = '<div class="cd"><p class="muted">불러오는 중…</p></div>';
    try{ data = await stmtData(st.stYm || ym); }catch(e){ box.innerHTML = `<div class="warn">${E(e.message||String(e))}</div>`; return; }
    cur ? detail() : list();
  };
  const payBar = () => { const w = S.workers.find(k => k.id===cur); if(!w) return ''; const pt = w.pay_type||'daily', amt = (pt==='weekly'?w.weekly_wage:w.monthly_wage)||'';
    return `<div class="cd"><h3>급여 유형 <span class="muted">(바꾸면 계정 설정에도 같이 적용)</span></h3><div class="fl">${cs('sp-pt', PAYLIST, pt)}</div>${pt!=='daily' ? `<label class="fl" style="margin-top:8px"><span>${pt==='weekly'?'주급':'월급'} 금액(원)</span><input type="number" min="0" step="10000" id="sp-wa" value="${amt}" placeholder="금액 입력"></label>` : ''}<div class="fl" style="margin-top:8px"><span>세금 방식</span>${cs('sp-tx', pt==='daily' ? TAXLIST : TAXLIST2, w.tax_type==='daily' ? 'none' : w.tax_type)}</div></div>`; };
  const bindPay = () => { const w = S.workers.find(k => k.id===cur); if(!w) return; const NO = '급여 유형용 SQL을 먼저 실행하세요 (young-re-paytype.sql)';
    const save = async (upd, msg) => { const r = await sb.from('workers').update(upd).eq('id',w.id); if(r.error){ Y.toast(/pay_type|wage|column/.test(r.error.message) ? NO : errMsg(r.error),1); detail(); return; } Object.assign(w, upd); await loadRefs(); Y.toast(msg||'저장했습니다'); await draw(); };
    const pt = $('#sp-pt'); if(pt) pt.onchange = () => { const v = pt.value || 'daily', upd = {pay_type:v}; if(v!=='daily' && w.tax_type==='daily') upd.tax_type = 'none'; save(upd); };
    const wa = $('#sp-wa'); if(wa) wa.onchange = () => { const v = wa.value.trim()==='' ? null : parseInt(wa.value,10); if(v!==null && !(v>=0)){ Y.toast('금액을 숫자로 입력하세요',1); return; } save(w.pay_type==='weekly' ? {weekly_wage:v} : {monthly_wage:v}); };
    const tx = $('#sp-tx'); if(tx) tx.onchange = () => save({tax_type:tx.value});
  };
  const optBar = () => `<div class="cd"><h3>표시 설정 <span class="muted">(명세서·이미지·엑셀에 공통 적용)</span></h3>
    <div class="fl"><span>보기 방식</span><div class="seg" id="so-mode"><button data-v="detail" class="${o.mode==='detail'?'on':''}">일자별 상세</button><button data-v="total" class="${o.mode==='total'?'on':''}">합계만</button></div></div>
    <div class="fl"><span>표시 항목</span><div class="cp" style="display:flex;flex-wrap:wrap;gap:6px">${[['loc','장소'],['time','근무시간'],['tax','세금'],['adj','가감 항목']].map(([k,n]) => `<button type="button" class="cb${o[k]?' on':''}" data-so="${k}" ${k!=='tax'&&k!=='adj'&&o.mode==='total'?'disabled':''}>${n}</button>`).join('')}</div></div></div>`;
  const bindOpt = () => {
    document.querySelectorAll('#so-mode button').forEach(b => b.onclick = () => { o.mode = b.dataset.v; ssave(o); cur ? detail() : list(); });
    document.querySelectorAll('[data-so]').forEach(b => b.onclick = () => { o[b.dataset.so] = !o[b.dataset.so]; ssave(o); cur ? detail() : list(); });
  };
  const list = () => {
    const L = data.list, box = $('#st-box');
    box.innerHTML = `<div class="cd bdr"><label class="fl"><span>정산 월</span><input type="month" id="st-ym" value="${st.stYm||ym}"></label><span style="flex:1"></span><button class="btn sm" id="st-xls">엑셀 전체 저장</button></div>` + optBar() +
      (data.adjMissing ? '<div class="warn">가감 항목을 쓰려면 SQL(young-re-adjust.sql)을 먼저 실행하세요. 지금은 가감 없이 표시됩니다.</div>' : '') +
      `<div class="cd"><h3>${ymLabel(st.stYm||ym)} 명세서 <span class="chip">${L.length}명</span></h3>${L.length ? `<div class="prl">${L.map((x,i) => { const m = stmtModel(x, st.stYm||ym, o), g = (S.workers.find(k => k.id===x.pay.worker_id)||{}).gender; return `<div class="pr"><div class="pn">${wn(x.pay.worker_name,g)}<span class="muted">${E(m.sub)}</span><span class="muted">${E(m.netLabel)} <b style="color:#2c3470">${money(m.net)}</b></span></div><button class="btn line sm" data-sv="${i}">보기</button><button class="btn line sm" data-si="${i}">이미지</button></div>`; }).join('')}</div>` : '<p class="muted">이 달의 기록이 없습니다</p>'}</div>`;
    bindOpt();
    $('#st-ym').onchange = e => { st.stYm = e.target.value; draw(); };
    box.querySelectorAll('[data-sv]').forEach(b => b.onclick = () => { cur = L[+b.dataset.sv].pay.worker_id; detail(); window.scrollTo(0,0); });
    box.querySelectorAll('[data-si]').forEach(b => b.onclick = () => stmtSaveImg(stmtModel(L[+b.dataset.si], st.stYm||ym, o)));
    $('#st-xls').onclick = () => {
      if(!L.length){ Y.toast('저장할 내용이 없습니다',1); return; }
      const wb = XLSX.utils.book_new(), used = new Set(), ms = L.map(x => stmtModel(x, st.stYm||ym, o));
      const head = ['이름','전화번호','급여 유형','세금 방식','근무일수','총 지급액']; if(o.adj) head.push('가감 합계'); if(o.tax) head.push('소득세','지방소득세','고용보험'); head.push(ms[0].netLabel);
      const sumRows = L.map((x,i) => { const p = x.pay, r = [p.worker_name, Y.fmtPhone(p.phone), ms[i].payType, ms[i].taxType, p.days, p.gross]; if(o.adj) r.push(x.adj.reduce((s,a) => s + a.amount, 0)); if(o.tax) r.push(p.income_tax||0, p.local_tax||0, p.emp_ins||0); r.push(ms[i].net); return r; });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ...sumRows]), '요약');
      ms.forEach(m => { let n = m.name.replace(/[\\\/\?\*\[\]:]/g,'').slice(0,28) || '시트', k = 2; while(used.has(n)) n = m.name.slice(0,26) + k++; used.add(n); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(stmtSheetRows(m)), n); });
      XLSX.writeFile(wb, '명세서_' + (st.stYm||ym) + '.xlsx');
    };
  };
  const detail = () => {
    const x = data.list.find(k => k.pay.worker_id===cur), box = $('#st-box'); if(!x){ cur = null; list(); return; }
    const m = stmtModel(x, st.stYm||ym, o);
    box.innerHTML = `<div class="cd bdr"><button class="btn line sm" id="sd-back">← 목록</button><span style="flex:1"></span><button class="btn line sm" id="sd-print">인쇄 / PDF</button><button class="btn sm" id="sd-img">이미지 저장</button></div>` + payBar() + optBar() +
      `<div class="cd">${stmtHtml(m)}</div>` +
      `<div class="cd"><h3>가감 항목 <span class="muted">(수당은 +, 공제는 −)</span></h3>${data.adjMissing ? '<p class="muted">young-re-adjust.sql을 먼저 실행해야 사용할 수 있습니다.</p>' : `${x.adj.length ? `<div class="prl">${x.adj.map(a => `<div class="pr"><div class="pn"><b>${E(a.label)}</b></div><b class="${a.amount<0?'neg':''}">${money(a.amount)}</b><button class="btn red sm" data-ad="${a.id}">삭제</button></div>`).join('')}</div>` : '<p class="muted">추가한 항목이 없습니다</p>'}
      <div class="row2" style="margin-top:12px"><label class="fl"><span>항목 이름</span><input id="ad-l" maxlength="20" placeholder="예: 식대, 선지급"></label><div class="fl"><span>구분</span>${cs('ad-s',[{id:'+',name:'지급 +'},{id:'-',name:'공제 −'}],'+')}</div><label class="fl"><span>금액(원)</span><input id="ad-a" type="number" min="0" step="1000"></label></div><div class="msg" id="ad-m"></div><button class="btn sm" id="ad-go" style="margin-top:6px">추가</button>
      <p class="muted" style="margin-top:10px">가감 항목은 세금 계산에 포함되지 않고 실지급액에만 더하거나 뺍니다.</p>`}</div>`;
    initChips(box); bindOpt(); bindPay();
    $('#sd-back').onclick = () => { cur = null; list(); };
    $('#sd-img').onclick = () => stmtSaveImg(m);
    $('#sd-print').onclick = () => { document.body.classList.add('printing'); window.print(); setTimeout(() => document.body.classList.remove('printing'), 500); };
    if(!data.adjMissing){
      box.querySelectorAll('[data-ad]').forEach(b => b.onclick = async () => { const r = await sb.from('pay_adjustments').delete().eq('id',b.dataset.ad); if(r.error){ Y.toast(errMsg(r.error),1); return; } await draw(); });
      $('#ad-go').onclick = async () => {
        const l = $('#ad-l').value.trim(), v = parseInt($('#ad-a').value,10), sgn = $('#ad-s').value === '-' ? -1 : 1, mm = $('#ad-m'); mm.textContent = '';
        if(!l){ mm.textContent = '항목 이름을 입력하세요'; return; } if(!(v > 0)){ mm.textContent = '금액을 입력하세요'; return; }
        const r = await sb.from('pay_adjustments').insert({worker_id:cur, ym:st.stYm||ym, label:l, amount:sgn*v}); if(r.error){ mm.textContent = errMsg(r.error); return; }
        await draw(); Y.toast('추가했습니다');
      };
    }
  };
  draw();
}

/* ---------- 항목·단가 ---------- */
async function pItems(){
  const today = Y.today();
  const render = () => {
    main.innerHTML = `<h1 class="pt">항목·단가</h1>
    <div class="cd"><h3>항목과 일당</h3><div class="icards">${S.items.map(i => { const h = S.prices.filter(p => p.item_id===i.id).slice(0,3), pr = priceOn(i.id,today); return `<div class="ic${i.active?'':' off'}"><div class="ich"><b>${E(i.name)}</b><span class="tag ${i.active?'on':'off'}">${i.active?'사용중':'중지'}</span></div><div class="icp">${pr ? Y.won(pr) : '<span class="chip wait">단가 미설정</span>'}</div><div class="muted" style="min-height:34px">${h.length ? h.map(p => (String(p.effective_from).startsWith('2000')?'기본':String(p.effective_from).slice(0,10)) + ' ' + Y.won(p.price)).join(' · ') : '단가 이력 없음'}</div><div class="acts"><button class="btn line sm" data-pr="${i.id}">단가 변경</button><button class="btn line sm" data-rn="${i.id}">이름</button><button class="btn line sm" data-tg="${i.id}">${i.active?'중지':'재개'}</button></div></div>`; }).join('')}</div>
    <div class="ih"><span>항목 추가</span></div><div class="row2"><label class="fl"><span>항목 이름</span><input id="i-name" maxlength="20" placeholder="예: 야간"></label><label class="fl"><span>일당(원)</span><input id="i-price" type="number" min="0" step="1000" placeholder="0"></label></div><div class="msg" id="i-msg"></div><button class="btn sm" id="i-add">항목 추가</button>
    <p class="muted" style="margin-top:10px">단가를 바꾸면 적용일부터 새 단가가 쓰이고, 이미 기록된 출퇴근의 일당은 그대로 유지됩니다.</p></div>
    <div class="cd"><h3>장소</h3><div class="chips" style="margin-bottom:12px">${S.locs.map(l => `<span class="chip ${l.active?'':'wait'}">${E(l.name)}${l.active?'':' (중지)'} <a href="#" data-lt="${l.id}" style="margin-left:6px;color:#4b5bd6;text-decoration:none">${l.active?'중지':'재개'}</a></span>`).join('')}</div><div class="row2"><label class="fl"><span>장소 이름</span><input id="l-name" maxlength="20" placeholder="예: 인천"></label></div><div class="msg" id="l-msg"></div><button class="btn sm" id="l-add">장소 추가</button></div>
    <div class="cd"><h3>세율·공제 설정</h3><div class="grid g2">${S.settings.map(s => `<label class="fl"><span>${E(s.label||s.key)}</span><input type="number" step="any" data-sk="${E(s.key)}" value="${s.value}"></label>`).join('')}</div><div class="msg" id="s-msg"></div><button class="btn sm" id="s-save" style="margin-top:6px">세율 저장</button><p class="muted" style="margin-top:10px">고용보험 적용 여부는 1(적용) 또는 0(미적용)으로 입력합니다. 법령이 바뀌면 여기서 수정하세요.</p></div>`;
    main.querySelectorAll('[data-pr]').forEach(b => b.onclick = () => { const it = S.items.find(x => x.id===b.dataset.pr);
      modal(it.name + ' 단가 변경', `<label class="fl"><span>새 일당(원)</span><input type="number" id="m-price" min="0" step="1000" value="${priceOn(it.id,today)||''}"></label><label class="fl"><span>적용일</span><input type="date" id="m-date" value="${today}"></label>`, async d => {
        const price = parseInt(d.querySelector('#m-price').value,10), day = d.querySelector('#m-date').value;
        if(isNaN(price) || price<0 || !day) return '일당과 적용일을 입력하세요';
        const r = await sb.from('item_prices').upsert({item_id:it.id,price,effective_from:day},{onConflict:'item_id,effective_from'});
        if(r.error) return errMsg(r.error); await loadRefs(); render(); Y.toast('단가를 저장했습니다'); return true; }); });
    main.querySelectorAll('[data-rn]').forEach(b => b.onclick = () => { const it = S.items.find(x => x.id===b.dataset.rn);
      modal('항목 이름 변경', `<label class="fl"><span>이름</span><input id="m-name" maxlength="20" value="${E(it.name)}"></label>`, async d => {
        const n = d.querySelector('#m-name').value.trim(); if(!n) return '이름을 입력하세요';
        const r = await sb.from('items').update({name:n}).eq('id',it.id); if(r.error) return errMsg(r.error);
        await loadRefs(); render(); Y.toast('변경했습니다'); return true; }); });
    main.querySelectorAll('[data-tg]').forEach(b => b.onclick = async () => { const it = S.items.find(x => x.id===b.dataset.tg); const r = await sb.from('items').update({active:!it.active}).eq('id',it.id); if(r.error) Y.toast(errMsg(r.error),1); else { await loadRefs(); render(); } });
    $('#i-add').onclick = async () => {
      const m = $('#i-msg'); m.textContent = ''; const n = $('#i-name').value.trim(), p = parseInt($('#i-price').value,10);
      if(!n){ m.textContent = '항목 이름을 입력하세요'; return; }
      const r = await sb.from('items').insert({name:n,sort:S.items.length+1}).select('id').single();
      if(r.error){ m.textContent = errMsg(r.error); return; }
      if(p > 0) await sb.from('item_prices').insert({item_id:r.data.id,price:p,effective_from:'2000-01-01'});
      await loadRefs(); render(); Y.toast('항목을 추가했습니다');
    };
    main.querySelectorAll('[data-lt]').forEach(a => a.onclick = async e => { e.preventDefault(); const l = S.locs.find(x => x.id===a.dataset.lt); const r = await sb.from('locations').update({active:!l.active}).eq('id',l.id); if(r.error) Y.toast(errMsg(r.error),1); else { await loadRefs(); render(); } });
    $('#l-add').onclick = async () => {
      const m = $('#l-msg'); m.textContent = ''; const n = $('#l-name').value.trim(); if(!n){ m.textContent = '장소 이름을 입력하세요'; return; }
      const r = await sb.from('locations').insert({name:n,sort:S.locs.length+1}); if(r.error){ m.textContent = errMsg(r.error); return; }
      await loadRefs(); render(); Y.toast('장소를 추가했습니다');
    };
    $('#s-save').onclick = async () => {
      const m = $('#s-msg'); m.textContent = '';
      const ups = [...main.querySelectorAll('[data-sk]')].map(i => ({key:i.dataset.sk,value:Number(i.value)}));
      if(ups.some(u => isNaN(u.value))){ m.textContent = '숫자만 입력하세요'; return; }
      const rs = await Promise.all(ups.map(u => sb.from('settings').update({value:u.value}).eq('key',u.key)));
      const bad = rs.find(x => x.error); if(bad){ m.textContent = errMsg(bad.error); return; }
      await loadRefs(); Y.toast('세율을 저장했습니다');
    };
  };
  render();
}

/* ---------- 계정 ---------- */
async function pAccounts(){
  const {data:profs,error} = await sb.from('profiles').select('*').order('created_at');
  if(error){ main.innerHTML = `<h1 class="pt">계정</h1><div class="warn">${E(error.message)}</div>`; return; }
  const pend = (profs||[]).filter(p => p.status==='pending'), staff = (profs||[]).filter(p => p.role==='staff' && p.status!=='pending'), admins = (profs||[]).filter(p => p.role==='admin' && p.status==='approved'), me = ((await sb.auth.getSession()).data.session||{}).user;
  const link = new URL('index.html', location.href).href;
  main.innerHTML = `<h1 class="pt">계정</h1>
  <div class="cd"><h3>직원 가입 링크</h3><div style="display:flex;gap:8px;flex-wrap:wrap"><input id="ln" readonly value="${E(link)}" style="flex:1;min-width:200px;height:42px;border-radius:12px;border:1px solid #e1e6fb;background:#f4f6ff;padding:0 12px;font:inherit;font-size:14px"><button class="btn sm" id="ln-cp">링크 복사</button></div><p class="muted" style="margin:8px 0 0">직원이 이 링크에서 가입하면 아래 승인 대기에 나타납니다.</p></div>
  <div class="cd"><h3>승인 대기 <span class="chip wait">${pend.length}명</span></h3>${pend.length ? `<div class="tw"><table class="t"><thead><tr><th>이름</th><th>전화번호</th><th>가입일</th><th></th></tr></thead><tbody>${pend.map(p => `<tr><td>${E(p.name)}</td><td>${Y.fmtPhone(p.phone)}</td><td>${String(p.created_at).slice(0,10)}</td><td><div class="acts"><button class="btn sm" data-ok="${p.id}">승인</button><button class="btn red sm" data-no="${p.id}">거절</button></div></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted">승인 대기 중인 가입이 없습니다</p>'}</div>
  <div class="cd"><h3>관리자 <span class="chip work">${admins.length}명</span></h3><div class="prl">${admins.map(p => `<div class="pr"><div class="pn"><b>${E(p.name)}</b><span class="muted">${Y.fmtPhone(p.phone)}</span></div>${me && p.id===me.id ? '<span class="chip">나</span>' : `<button class="btn red sm" data-rv="${p.id}">관리자 해제</button>`}</div>`).join('')}</div></div>
  <div class="cd"><h3>직원 계정</h3>${staff.length ? `<div class="tw"><table class="t"><thead><tr><th>이름</th><th>전화번호</th><th>상태</th><th></th></tr></thead><tbody>${staff.map(p => `<tr><td>${E(p.name)}</td><td>${Y.fmtPhone(p.phone)}</td><td><span class="tag ${p.status==='approved'?'on':'off'}">${p.status==='approved'?'승인됨':'거절됨'}</span></td><td>${p.status==='approved' ? `<div class="acts"><button class="btn line sm" data-gr="${p.id}">관리자로 지정</button><button class="btn red sm" data-no="${p.id}">사용 중지</button></div>` : `<button class="btn sm" data-ok="${p.id}">승인</button>`}</td></tr>`).join('')}</tbody></table></div>` : '<p class="muted">직원 계정이 없습니다</p>'}</div>
  <div class="cd"><div class="ih"><h3 style="margin:0">근로자 <span class="chip">${S.workers.length}명</span></h3><button class="btn sm" id="w-add">근로자 등록</button></div><p class="muted" style="margin:8px 0 10px">QR로 출퇴근한 알바도 자동으로 등록되고, 여기서 미리 등록할 수도 있습니다. 성별과 세금 방식은 카드에서 고르세요.</p>${S.workers.length ? `<div class="icards wk">${S.workers.map(w => `<div class="ic"><div class="ich"><b class="gn ${gcls(w.gender)}">${E(w.name)}</b><span class="tag">${w.is_staff?'직원':'알바'}</span></div><div class="muted" style="margin:2px 0 8px">${Y.fmtPhone(w.phone)}</div><div class="muted" style="margin-bottom:4px">성별</div>${cs('gd-'+w.id, GLIST, w.gender||'')}<div class="muted" style="margin:10px 0 4px">급여 유형</div>${cs('pt-'+w.id, PAYLIST, w.pay_type||'daily')}${(w.pay_type||'daily')!=='daily' ? `<label class="fl" style="margin-top:8px"><span>${w.pay_type==='weekly'?'주급':'월급'} 금액(원)</span><input type="number" min="0" step="10000" data-wa="${w.id}" value="${(w.pay_type==='weekly'?w.weekly_wage:w.monthly_wage)||''}" placeholder="금액 입력"></label>` : ''}<div class="muted" style="margin:10px 0 4px">세금 방식</div>${cs('tx-'+w.id, (w.pay_type||'daily')==='daily' ? TAXLIST : TAXLIST2, w.tax_type)}${(() => { const seen = new Set(), t = Y.today(); const lines = (S.wprices||[]).filter(p => p.worker_id===w.id && String(p.effective_from)<=t && !seen.has(p.item_id) && seen.add(p.item_id)).map(p => `${E((S.items.find(i => i.id===p.item_id)||{}).name||'')} ${Y.won(p.price)}`); return lines.length ? `<div class="muted" style="margin-top:10px">개인 단가: ${lines.join(' · ')}</div>` : ''; })()}<input data-mm="${w.id}" value="${E(w.memo||'')}" maxlength="60" placeholder="메모" style="margin-top:10px;width:100%;height:38px;border-radius:10px;border:1px solid #e1e6fb;background:#f4f6ff;padding:0 10px;font:inherit;font-size:14px"><div class="acts" style="margin-top:10px">${(w.pay_type||'daily')==='daily' ? `<button class="btn line sm" data-wp="${w.id}">개인 단가</button>` : ''}<button class="btn line sm" data-wed="${w.id}">이름·번호 수정</button><button class="btn red sm" data-wdel="${w.id}">삭제</button></div></div>`).join('')}</div>` : '<p class="muted">등록된 근로자가 없습니다</p>'}</div>`;
  $('#ln-cp').onclick = async () => { try{ await navigator.clipboard.writeText(link); Y.toast('링크를 복사했습니다'); }catch(e){ $('#ln').select(); } };
  const setStatus = async (id, s) => { const r = await sb.from('profiles').update({status:s}).eq('id',id); if(r.error){ Y.toast(errMsg(r.error),1); return; } await loadRefs(); Y.toast(s==='approved'?'승인했습니다':'처리했습니다'); pAccounts(); };
  const setRole = async (id, role) => { const r = await sb.from('profiles').update({role, status:'approved'}).eq('id',id); if(r.error){ Y.toast(errMsg(r.error),1); return; } Y.toast(role==='admin'?'관리자로 지정했습니다':'관리자 권한을 해제했습니다'); pAccounts(); };
  main.querySelectorAll('[data-gr]').forEach(b => b.onclick = () => { if(confirm('이 직원에게 관리자 권한을 줄까요?\n금액·정산을 포함한 모든 화면을 볼 수 있게 됩니다.')) setRole(b.dataset.gr,'admin'); });
  main.querySelectorAll('[data-rv]').forEach(b => b.onclick = () => { if(confirm('관리자 권한을 해제할까요?\n직원 계정으로 바뀝니다.')) setRole(b.dataset.rv,'staff'); });
  main.querySelectorAll('[data-ok]').forEach(b => b.onclick = () => setStatus(b.dataset.ok,'approved'));
  main.querySelectorAll('[data-no]').forEach(b => b.onclick = () => { if(confirm('이 계정을 사용할 수 없게 할까요?')) setStatus(b.dataset.no,'rejected'); });
  S.workers.forEach(w => { const g = $('#gd-'+w.id); if(g) g.onchange = async () => { const r = await sb.from('workers').update({gender:g.value||null}).eq('id',w.id); if(r.error) Y.toast(errMsg(r.error),1); else { w.gender = g.value||null; const b = g.closest('.ic').querySelector('.ich b'); b.className = 'gn ' + gcls(w.gender); Y.toast('저장했습니다'); } }; });
  const NOPT = '급여 유형용 SQL을 먼저 실행하세요 (young-re-paytype.sql)';
  S.workers.forEach(w => { const el = $('#pt-'+w.id); if(el) el.onchange = async () => {
    const v = el.value || 'daily', upd = {pay_type:v}; let note = '';
    if(v!=='daily' && w.tax_type==='daily'){ upd.tax_type = 'none'; note = ' 세금 방식은 없음으로 바꿨습니다'; }
    const r = await sb.from('workers').update(upd).eq('id',w.id);
    if(r.error){ Y.toast(/pay_type|column/.test(r.error.message) ? NOPT : errMsg(r.error),1); pAccounts(); return; }
    await loadRefs(); Y.toast('저장했습니다.' + note); pAccounts(); }; });
  main.querySelectorAll('[data-wa]').forEach(inp => inp.onchange = async () => { const w = S.workers.find(x => x.id===inp.dataset.wa), v = inp.value.trim() === '' ? null : parseInt(inp.value,10);
    if(v !== null && !(v >= 0)){ Y.toast('금액을 숫자로 입력하세요',1); return; }
    const r = await sb.from('workers').update(w.pay_type==='weekly' ? {weekly_wage:v} : {monthly_wage:v}).eq('id',w.id);
    if(r.error){ Y.toast(/wage|column/.test(r.error.message) ? NOPT : errMsg(r.error),1); return; } if(w.pay_type==='weekly') w.weekly_wage = v; else w.monthly_wage = v; Y.toast('저장했습니다'); });
  S.workers.forEach(w => { const el = $('#tx-'+w.id); if(el) el.onchange = async () => { const r = await sb.from('workers').update({tax_type:el.value}).eq('id',w.id); if(r.error) Y.toast(errMsg(r.error),1); else { w.tax_type = el.value; Y.toast('저장했습니다'); } }; });
  $('#w-add').onclick = () => modal('근로자 등록', `<label class="fl"><span>이름</span><input id="n-nm" maxlength="20"></label><label class="fl"><span>전화번호</span><input id="n-ph" type="tel" inputmode="numeric" placeholder="010-0000-0000"></label><div class="fl"><span>성별</span>${cs('n-g', GLIST, '')}</div><div class="fl"><span>세금 방식</span>${cs('n-t', TAXLIST, 'daily')}</div>`, async d => {
    const n = d.querySelector('#n-nm').value.trim(), p = Y.digits(d.querySelector('#n-ph').value);
    if(n.length<2) return '이름을 입력하세요'; if(!Y.validPhone(p)) return '전화번호를 정확히 입력하세요';
    if(S.workers.some(x => x.phone===p)) return '이미 등록된 전화번호입니다';
    const r = await sb.from('workers').insert({name:n,phone:p,gender:d.querySelector('#n-g').value||null,tax_type:d.querySelector('#n-t').value||'daily'}); if(r.error) return errMsg(r.error);
    await loadRefs(); Y.toast('등록했습니다'); pAccounts(); return true; }, '등록');
  main.querySelectorAll('[data-wp]').forEach(b => b.onclick = () => { const w = S.workers.find(x => x.id===b.dataset.wp), t = Y.today(), its = S.items.filter(i => i.active);
    modal(w.name + ' 개인 단가', `<p class="muted" style="margin:0 0 10px">같은 항목이라도 이 사람에게만 다른 일당을 줄 때 입력합니다. 비워 두면 항목 기본 단가가 적용됩니다.</p>
    <label class="fl"><span>적용 시작일</span><input type="date" id="wp-d" value="${t}"></label>
    ${its.map(i => { const cur = (S.wprices||[]).find(p => p.worker_id===w.id && p.item_id===i.id && String(p.effective_from)<=t); return `<div class="pr" style="margin-bottom:8px"><div class="pn"><b>${E(i.name)}</b><span class="muted">기본 ${Y.won(priceOn(i.id,t))}${cur ? ' · 현재 개인 ' + Y.won(cur.price) : ''}</span></div><input type="number" min="0" step="1000" data-wi="${i.id}" value="${cur ? cur.price : ''}" placeholder="기본" style="width:120px;height:40px;border-radius:10px;border:1px solid #e1e6fb;background:#f4f6ff;padding:0 10px;font:inherit;text-align:right"></div>`; }).join('')}
    <label class="fl" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="wp-back"><span>적용 시작일 이후 이미 기록된 출근 건도 이 금액으로 바꾸기</span></label>`, async d => {
      const day = d.querySelector('#wp-d').value; if(!day) return '적용 시작일을 선택하세요';
      const back = d.querySelector('#wp-back').checked;
      for(const inp of d.querySelectorAll('[data-wi]')){
        const item = inp.dataset.wi, v = inp.value.trim(), cur = (S.wprices||[]).find(p => p.worker_id===w.id && p.item_id===item && String(p.effective_from)<=day);
        if(v === ''){ if(cur){ const r = await sb.from('worker_prices').delete().eq('worker_id',w.id).eq('item_id',item); if(r.error) return errMsg(r.error); if(back){ await sb.from('attendance').update({price_applied:priceOn(item,day)}).eq('worker_id',w.id).eq('item_id',item).gte('work_date',day); } } continue; }
        const p = parseInt(v,10); if(!(p >= 0)) return '금액을 숫자로 입력하세요';
        if(cur && cur.price===p && String(cur.effective_from)===day && !back) continue;
        const r = await sb.from('worker_prices').upsert({worker_id:w.id,item_id:item,price:p,effective_from:day},{onConflict:'worker_id,item_id,effective_from'}); if(r.error) return /worker_prices/.test(r.error.message) ? '개인 단가용 SQL을 먼저 실행하세요 (young-re-workerprice.sql)' : errMsg(r.error);
        if(back){ const u = await sb.from('attendance').update({price_applied:p}).eq('worker_id',w.id).eq('item_id',item).gte('work_date',day); if(u.error) return errMsg(u.error); }
      }
      await loadRefs(); Y.toast('개인 단가를 저장했습니다'); pAccounts(); return true; }, '저장'); });
  main.querySelectorAll('[data-wed]').forEach(b => b.onclick = () => { const w = S.workers.find(x => x.id===b.dataset.wed); modal('근로자 정보 수정', `<label class="fl"><span>이름</span><input id="w-n" value="${E(w.name)}" maxlength="20"></label><label class="fl"><span>전화번호</span><input id="w-p" inputmode="numeric" value="${E(Y.fmtPhone(w.phone))}" ${w.is_staff?'readonly':''}></label>${w.is_staff?'<p class="muted">직원의 전화번호는 로그인 아이디라 여기서 바꿀 수 없습니다. 이름만 수정됩니다.</p>':''}`, async d => {
    const n = d.querySelector('#w-n').value.trim(), p = Y.digits(d.querySelector('#w-p').value);
    if(!n) return '이름을 입력하세요'; if(!w.is_staff && !Y.validPhone(p)) return '전화번호를 정확히 입력하세요';
    const upd = {name:n}; if(!w.is_staff) upd.phone = p;
    const r = await sb.from('workers').update(upd).eq('id',w.id);
    if(r.error) return /duplicate|unique/i.test(r.error.message) ? '이미 등록된 전화번호입니다' : errMsg(r.error);
    await loadRefs(); Y.toast('수정했습니다'); pAccounts(); return true; }); });
  main.querySelectorAll('[data-wdel]').forEach(b => b.onclick = async () => { const w = S.workers.find(x => x.id===b.dataset.wdel);
    if(w.is_staff){ Y.toast('직원은 위 직원 계정에서 사용 중지하세요',1); return; }
    const c = await sb.from('attendance').select('id',{count:'exact',head:true}).eq('worker_id',w.id);
    if(c.count>0){ Y.toast(`출퇴근 기록이 ${c.count}건 있어 삭제할 수 없습니다`,1); return; }
    if(!confirm(w.name+' 님을 삭제할까요?')) return;
    const r = await sb.from('workers').delete().eq('id',w.id); if(r.error){ Y.toast(errMsg(r.error),1); return; }
    await loadRefs(); Y.toast('삭제했습니다'); pAccounts(); });
  main.querySelectorAll('[data-mm]').forEach(i => i.onchange = async () => { const r = await sb.from('workers').update({memo:i.value.trim()||null}).eq('id',i.dataset.mm); if(r.error) Y.toast(errMsg(r.error),1); else Y.toast('저장했습니다'); });
}


/* ---------- 초기화 ---------- */
const RESETS = [
  {k:'att',  name:'출퇴근 기록', tbl:'attendance',  desc:'모든 출근·퇴근 기록과 일당 정산 내역', run: async () => await sb.from('attendance').delete().not('id','is',null)},
  {k:'asg',  name:'업무 배치',   tbl:'assignments', desc:'업무 부여로 배치한 내역', run: async () => await sb.from('assignments').delete().not('id','is',null)},
  {k:'qr',   name:'QR 코드',     tbl:'qr_codes',    desc:'만들어 둔 모든 QR (이미 인쇄한 QR은 더 이상 쓸 수 없음)', run: async () => await sb.from('qr_codes').delete().not('id','is',null)},
  {k:'note', name:'작업내역',    tbl:'work_notes',  desc:'항목·날짜별 작업내역 카드', run: async () => await sb.from('work_notes').delete().not('id','is',null)},
  {k:'alba', name:'알바 명단',   tbl:'workers',     filter:q => q.eq('is_staff',false), desc:'알바 근로자 명단. 그 사람들의 출퇴근·배치 기록도 함께 삭제됨', run: async () => await sb.from('workers').delete().eq('is_staff',false)}
];
async function pReset(){
  main.innerHTML = '<h1 class="pt">초기화</h1><div class="cd"><p class="muted">불러오는 중…</p></div>';
  const cnt = await Promise.all(RESETS.map(async r => { let q = sb.from(r.tbl).select('id',{count:'exact',head:true}); if(r.filter) q = r.filter(q); const x = await q; return x.count ?? 0; }));
  const total = cnt.reduce((a,b) => a + b, 0);
  main.innerHTML = `<h1 class="pt">초기화</h1>
  <div class="cd"><h3>항목별 초기화</h3><p class="muted" style="margin:0 0 12px">지운 데이터는 되돌릴 수 없습니다. 지우기 전에 조회 화면에서 엑셀로 저장해 두세요. 항목·단가, 장소, 세금 설정, 직원 계정은 어느 초기화에서도 지워지지 않습니다.</p>
  <div class="prl">${RESETS.map((r,i) => `<div class="pr"><div class="pn"><b>${r.name}</b><span class="muted">${r.desc}</span></div><span class="chip ${cnt[i]?'work':''}">${cnt[i]}건</span><button class="btn red sm" data-rs="${i}" ${cnt[i]?'':'disabled'}>초기화</button></div>`).join('')}</div></div>
  <div class="cd" style="border:1.5px solid #f3c6cf"><h3>전체 초기화</h3><p class="muted" style="margin:0 0 12px">위 5가지(출퇴근 기록, 업무 배치, QR, 작업내역, 알바 명단)를 한 번에 지웁니다. 새로 시작할 때 사용하세요.</p><button class="btn red" id="rs-all" ${total?'':'disabled'}>전체 초기화</button></div>`;
  const confirmReset = (title, text, run) => modal(title, `<p style="line-height:1.7;margin:0 0 12px">${text}</p><label class="fl"><span>계속하려면 <b>초기화</b>를 입력하세요</span><input id="rs-t" autocomplete="off"></label>`, async d => {
    if(d.querySelector('#rs-t').value.trim() !== '초기화') return '"초기화"를 정확히 입력하세요';
    const r = await run(); if(r && r.error) return '지우지 못했습니다: ' + errMsg(r.error);
    await loadRefs(); Y.toast('초기화했습니다'); pReset(); return true; }, '초기화');
  main.querySelectorAll('[data-rs]').forEach(b => b.onclick = () => { const r = RESETS[+b.dataset.rs]; confirmReset(r.name + ' 초기화', `<b>${r.name}</b> ${cnt[+b.dataset.rs]}건을 모두 지웁니다.<br>${r.desc}.`, r.run); });
  $('#rs-all').onclick = () => confirmReset('전체 초기화', `출퇴근 기록, 업무 배치, QR, 작업내역, 알바 명단 <b>${total}건</b>을 모두 지웁니다.`, async () => { for(const r of RESETS){ const x = await r.run(); if(x && x.error) return x; } return null; });
}

/* ---------- 라우팅 ---------- */
const PAGES = {board:pBoard,assign:pAssign,qr:pQr,notes:pNotes,query:pQuery,pay:pPay,stmt:pStmt,items:pItems,accounts:pAccounts,reset:pReset};
async function route(){
  if(rt){ sb.removeChannel(rt); rt = null; }
  const p = PAGES[(location.hash||'#board').slice(1)] ? (location.hash||'#board').slice(1) : 'board';
  $('#tabs').innerHTML = Object.entries(NAMES).map(([k,v]) => `<a href="#${k}" class="${k===p?'on':''}"><span class="dot">${v[0]}</span>${v}</a>`).join('');
  if(window.innerWidth < 1000) document.body.classList.remove('side-open');
  try{ await PAGES[p](); }catch(e){ main.innerHTML = `<div class="warn">화면을 그리지 못했습니다: ${E(e.message)}</div>`; }
  window.scrollTo(0,0);
}
(async () => {
  const me = await Y.guard('admin'); if(!me) return;
  $('#who').textContent = me.name + ' 관리자';
  $('#lo').onclick = Y.logout;
  const tg = () => { document.body.classList.toggle('side-open'); try{ if(innerWidth>=1000) localStorage.setItem('yre_side', document.body.classList.contains('side-open')?'1':'0'); }catch(e){} };
  let sx=0, sy=0;
  document.addEventListener('touchstart', e => { sx=e.touches[0].clientX; sy=e.touches[0].clientY; }, {passive:true});
  document.addEventListener('touchend', e => { if(innerWidth>=1000) return; const t=e.changedTouches[0], dx=t.clientX-sx, dy=Math.abs(t.clientY-sy); if(dy>60 || Math.abs(dx)<60) return; const open=document.body.classList.contains('side-open'); if(dx>0 && !open && sx<40) document.body.classList.add('side-open'); else if(dx<0 && open) document.body.classList.remove('side-open'); }, {passive:true});
  $('#hb').onclick = tg; $('#ov').onclick = () => document.body.classList.remove('side-open');
  try{ if(innerWidth>=1000 && localStorage.getItem('yre_side')!=='0') document.body.classList.add('side-open'); }catch(e){}
  await loadRefs();
  window.addEventListener('hashchange', route);
  route();
})();
})();
