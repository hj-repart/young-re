(function(){
const Y = window.YRE, sb = Y.sb, E = Y.esc;
const $ = s => document.querySelector(s);
const main = $('#main');
const NAMES = {board:'상황판',assign:'업무 부여',qr:'QR 만들기',notes:'작업내역',query:'조회',pay:'정산',items:'항목·단가',accounts:'계정'};
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
    sb.from('settings').select('*').order('key')]);
  const bad = r.find(x => x.error);
  if(bad) Y.toast('데이터를 불러오지 못했습니다: ' + bad.error.message, 1);
  S.items = r[0].data||[]; S.locs = r[1].data||[]; S.workers = r[2].data||[]; S.prices = r[3].data||[]; S.settings = r[4].data||[];
}
const activeItems = () => S.items.filter(i => i.active);
const activeLocs = () => S.locs.filter(l => l.active);
const priceOn = (id,d) => { const p = S.prices.find(p => p.item_id===id && String(p.effective_from)<=d); return p ? p.price : 0; };
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
  return sb.from('attendance').insert({worker_id:w.worker_id,work_date:w.work_date,item_id:w.item_id,location_id:w.location_id||null,assignment_id:w.assignment_id||null,check_in:w.check_in||nowFor(w.work_date),price_applied:priceOn(w.item_id,w.work_date),source:'admin'});
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

/* ---------- 업무 부여 ---------- */
async function pAssign(){
  const staff = S.workers;
  main.innerHTML = `<h1 class="pt">업무 부여</h1><div class="cd"><h3>직원에게 업무 부여</h3>
    <div class="row2"><label class="fl"><span>날짜</span><input type="date" id="a-date" value="${st.aDate||Y.today()}"></label>
    <div class="fl"><span>항목</span>${cs('a-item', activeItems(), st.aItem)}</div>
    <div class="fl"><span>장소</span>${cs('a-loc', activeLocs(), st.aLoc, '선택 안 함')}</div></div>
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
    $('#a-list').querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('이 업무를 삭제할까요?')) return; const r = await sb.from('assignments').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast('삭제했습니다'); list(); } });
  };
  $('#a-item').onchange = () => { st.aItem = $('#a-item').value; cards(); };
  $('#a-date').onchange = list;
  let mode = 'plan';
  document.querySelectorAll('#a-mode button').forEach(b => b.onclick = () => { mode = b.dataset.m; document.querySelectorAll('#a-mode button').forEach(x => x.classList.toggle('on', x===b)); $('#a-mh').textContent = mode==='in' ? '선택한 사람을 관리자가 바로 출근 처리합니다. (오늘이면 지금 시각, 다른 날짜면 09:00 기준)' : '직원이 앱에서 직접 출근 버튼을 누릅니다.'; });
  const syncPick = () => document.querySelectorAll('.wp input').forEach(i => i.closest('.wp').classList.toggle('on', i.checked));
  document.querySelectorAll('.wp input').forEach(i => i.onchange = syncPick);
  const all = $('#a-all'); if(all) all.onclick = () => { document.querySelectorAll('#a-staff input[data-st="1"]').forEach(c => c.checked = true); syncPick(); };
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
    $('#a-go').disabled = true;
    const {data:asg,error} = await sb.from('assignments').upsert(ids.map(w => ({work_date:date,worker_id:w,item_id:item,location_id:loc,task_text:task||null})), {onConflict:'work_date,worker_id,item_id'}).select('id,worker_id');
    if(error){ $('#a-go').disabled = false; m.textContent = errMsg(error); return; }
    let done = 0;
    if(mode==='in'){
      const rows = (asg||[]).map(a => ({worker_id:a.worker_id,work_date:date,item_id:item,location_id:loc,assignment_id:a.id,check_in:nowFor(date),price_applied:priceOn(item,date),source:'admin'}));
      const x = await sb.from('attendance').upsert(rows,{onConflict:'worker_id,work_date,item_id',ignoreDuplicates:true}).select('id');
      if(x.error){ $('#a-go').disabled = false; m.textContent = '배치는 저장했지만 출근 처리에 실패했습니다: ' + errMsg(x.error); list(); return; }
      done = (x.data||[]).length;
    }
    $('#a-go').disabled = false;
    await saveNote(date, item, loc, task);
    st.aItem = item; Y.toast(mode==='in' ? `${ids.length}명 배치, ${done}명 출근 처리했습니다` + (done<ids.length ? ` (${ids.length-done}명은 이미 출근 기록이 있습니다)` : '') : ids.length + '명에게 업무를 부여했습니다'); document.querySelectorAll('#a-staff input').forEach(c => c.checked = false); syncPick(); list(); cards();
  };
  cards(); list();
}

/* ---------- QR 만들기 ---------- */
function qrUrl(token){ return new URL('qr.html?t=' + token, location.href).href; }
async function qrImage(token, lines){
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
    <div class="fl"><span>항목</span>${cs('q-item', activeItems(), st.qItem)}</div>
    <div class="fl"><span>장소</span>${cs('q-loc', activeLocs(), st.qLoc, '선택 안 함')}</div></div>
    <div class="ih"><span>작업 내용</span></div><div class="ncs" id="q-cards"></div>
    <label class="fl"><textarea id="q-task" placeholder="예: 신한 현장 정리"></textarea></label>
    <div class="msg" id="q-msg"></div><button class="btn sm" id="q-go" style="margin-top:6px">QR 만들기</button></div>
    <div class="cd"><h3>QR 목록 <span class="muted" id="q-day"></span></h3><label class="fl" style="max-width:200px;margin-bottom:10px"><span>조회 날짜</span><input type="date" id="q-fd" value="${st.qDate||Y.today()}"></label><div class="qrs" id="q-list"></div></div>`;
  const cards = () => noteCards($('#q-cards'), $('#q-item').value, n => { $('#q-task').value = n.content; if(n.location_id && S.locs.find(l=>l.id===n.location_id&&l.active)) $('#q-loc').value = n.location_id; });
  const list = async () => {
    const d = $('#q-fd').value; $('#q-day').textContent = d ? Y.fmtDate(d) : '';
    const {data,error} = await sb.from('qr_codes').select('id,token,work_date,task_text,active,items(name),locations(name)').eq('work_date',d).order('created_at',{ascending:false});
    if(error){ $('#q-list').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    if(!(data||[]).length){ $('#q-list').innerHTML = '<p class="muted">이 날짜에 만든 QR이 없습니다</p>'; return; }
    const imgs = await Promise.all(data.map(r => qrImage(r.token, [(r.items||{}).name||'', (r.locations||{}).name||'장소 미지정', Y.fmtDate(r.work_date)])));
    $('#q-list').innerHTML = data.map((r,i) => `<div class="qc${r.active?'':' off'}"><img src="${imgs[i]}" alt="QR"><div style="margin:8px 0 2px"><b>${E((r.items||{}).name)}</b> · ${E((r.locations||{}).name||'-')} <span class="tag ${r.active?'on':'off'}">${r.active?'사용중':'중지'}</span></div><div class="muted" style="min-height:18px">${E(r.task_text||'')}</div><div class="acts" style="justify-content:center;margin-top:8px"><a class="btn line sm" style="display:inline-flex;align-items:center;text-decoration:none" download="QR_${r.work_date}_${E((r.items||{}).name)}_${E((r.locations||{}).name||'')}.png" href="${imgs[i]}">이미지 저장</a><button class="btn line sm" data-cp="${r.token}">링크 복사</button><button class="btn line sm" data-tg="${r.id}" data-on="${r.active}">${r.active?'사용 중지':'다시 사용'}</button><button class="btn red sm" data-del="${r.id}">삭제</button></div></div>`).join('');
    const box = $('#q-list');
    box.querySelectorAll('[data-cp]').forEach(b => b.onclick = async () => { try{ await navigator.clipboard.writeText(qrUrl(b.dataset.cp)); Y.toast('링크를 복사했습니다'); }catch(e){ prompt('링크를 복사하세요', qrUrl(b.dataset.cp)); } });
    box.querySelectorAll('[data-tg]').forEach(b => b.onclick = async () => { const r = await sb.from('qr_codes').update({active: b.dataset.on!=='true'}).eq('id',b.dataset.tg); if(r.error) Y.toast(errMsg(r.error),1); else list(); });
    box.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('이 QR을 삭제할까요?\n이미 찍힌 출퇴근 기록은 남습니다.')) return; const r = await sb.from('qr_codes').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast('삭제했습니다'); list(); } });
  };
  $('#q-item').onchange = () => { st.qItem = $('#q-item').value; cards(); };
  $('#q-fd').onchange = list;
  $('#q-go').onclick = async () => {
    const m = $('#q-msg'); m.textContent = '';
    const date = $('#q-date').value, item = $('#q-item').value, loc = $('#q-loc').value || null, task = $('#q-task').value.trim();
    if(!date || !item){ m.textContent = '날짜와 항목을 선택하세요'; return; }
    $('#q-go').disabled = true;
    const {error} = await sb.from('qr_codes').insert({work_date:date,item_id:item,location_id:loc,task_text:task||null});
    $('#q-go').disabled = false;
    if(error){ m.textContent = errMsg(error); return; }
    await saveNote(date, item, loc, task);
    st.qDate = date; st.qItem = item; $('#q-fd').value = date; Y.toast('QR을 만들었습니다'); list(); cards();
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
    <div class="fl"><span>사람</span>${cs('f-worker', S.workers, q.worker, '전체')}</div>
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
    $('#f-res').innerHTML = `<div class="tw"><table class="t"><thead><tr><th>일자</th><th>이름</th><th>항목</th><th>장소</th><th>출근</th><th>퇴근</th><th>근무시간</th><th class="r">일당</th><th></th></tr></thead><tbody>${keys.map(k => { const g = groups.get(k), sum = g.reduce((s,r) => s + (r.price_applied||0), 0); return `<tr class="gh"><td colspan="7">${E(k)} · ${g.length}건</td><td class="r">${Y.won(sum)}</td><td></td></tr>` + g.map(r => `<tr><td>${Y.fmtDate(r.work_date)}</td><td>${wn((r.workers||{}).name,(r.workers||{}).gender)}</td><td>${E((r.items||{}).name)}</td><td>${E((r.locations||{}).name||'-')}</td><td>${Y.fmtTime(r.check_in)}</td><td>${r.check_out?Y.fmtTime(r.check_out):'<span class="chip work">근무중</span>'}</td><td>${Y.fmtDur(r.check_in,r.check_out)}</td><td class="r">${Y.won(r.price_applied)}</td><td><div class="acts"><button class="btn line sm" data-ed="${r.id}">수정</button><button class="btn red sm" data-del="${r.id}">삭제</button></div></td></tr>`).join(''); }).join('')}<tr class="tot"><td colspan="7">합계 · ${rows.length}건</td><td class="r">${Y.won(total)}</td><td></td></tr></tbody></table></div>`;
    $('#f-res').querySelectorAll('[data-ed]').forEach(b => b.onclick = () => edit(rows.find(r => r.id===b.dataset.ed)));
    $('#f-res').querySelectorAll('[data-del]').forEach(b => b.onclick = async () => { if(!confirm('이 출퇴근 기록을 삭제할까요?')) return; const r = await sb.from('attendance').delete().eq('id',b.dataset.del); if(r.error) Y.toast(errMsg(r.error),1); else { Y.toast('삭제했습니다'); load(); } });
  };
  const load = async () => {
    q.from = $('#f-from').value; q.to = $('#f-to').value; q.worker = $('#f-worker').value; q.item = $('#f-item').value; q.loc = $('#f-loc').value;
    let b = sb.from('attendance').select('id,work_date,check_in,check_out,price_applied,note,worker_id,item_id,location_id,workers(name,phone,gender),items(name),locations(name)').gte('work_date',q.from).lte('work_date',q.to).order('work_date').order('check_in').limit(5000);
    if(q.worker) b = b.eq('worker_id',q.worker); if(q.item) b = b.eq('item_id',q.item); if(q.loc) b = b.eq('location_id',q.loc);
    const {data,error} = await b;
    if(error){ $('#f-res').innerHTML = `<div class="warn">${E(error.message)}</div>`; return; }
    rows = data||[]; draw();
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
      const price = g('#e-price') !== '' && parseInt(g('#e-price'),10) > 0 ? parseInt(g('#e-price'),10) : priceOn(item, day);
      const r = await sb.from('attendance').insert({worker_id:w,work_date:day,item_id:item,location_id:g('#e-loc')||null,check_in:cin,check_out:cout,price_applied:price,source:'admin',note:g('#e-note')||null});
      if(r.error) return errMsg(r.error); Y.toast('추가했습니다'); load(); return true;
    }, '추가');
  };
  document.querySelectorAll('#f-view button').forEach(b => b.onclick = () => { q.view = b.dataset.v; draw(); });
  ['f-from','f-to','f-worker','f-item','f-loc'].forEach(id => $('#'+id).onchange = load);
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
      sb.from('v_month_pay').select('*').eq('ym',v).order('worker_name'),
      sb.from('attendance').select('id',{count:'exact',head:true}).eq('price_applied',0).gte('work_date',v+'-01').lte('work_date',v+'-31')]);
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
  <div class="cd"><h3>근로자 세금 방식</h3><p class="muted" style="margin:0 0 10px">QR로 출퇴근한 알바도 여기에 자동으로 등록됩니다. 정산에 쓸 세금 방식을 근로자별로 고르세요.</p>${S.workers.length ? `<div class="icards wk">${S.workers.map(w => `<div class="ic"><div class="ich"><b class="gn ${gcls(w.gender)}">${E(w.name)}</b><span class="tag">${w.is_staff?'직원':'알바'}</span></div><div class="muted" style="margin:2px 0 8px">${Y.fmtPhone(w.phone)}</div><div class="muted" style="margin-bottom:4px">성별</div>${cs('gd-'+w.id, GLIST, w.gender||'')}<div class="muted" style="margin:10px 0 4px">세금 방식</div>${cs('tx-'+w.id, TAXLIST, w.tax_type)}<input data-mm="${w.id}" value="${E(w.memo||'')}" maxlength="60" placeholder="메모" style="margin-top:10px;width:100%;height:38px;border-radius:10px;border:1px solid #e1e6fb;background:#f4f6ff;padding:0 10px;font:inherit;font-size:14px"></div>`).join('')}</div>` : '<p class="muted">등록된 근로자가 없습니다</p>'}</div>`;
  $('#ln-cp').onclick = async () => { try{ await navigator.clipboard.writeText(link); Y.toast('링크를 복사했습니다'); }catch(e){ $('#ln').select(); } };
  const setStatus = async (id, s) => { const r = await sb.from('profiles').update({status:s}).eq('id',id); if(r.error){ Y.toast(errMsg(r.error),1); return; } await loadRefs(); Y.toast(s==='approved'?'승인했습니다':'처리했습니다'); pAccounts(); };
  const setRole = async (id, role) => { const r = await sb.from('profiles').update({role, status:'approved'}).eq('id',id); if(r.error){ Y.toast(errMsg(r.error),1); return; } Y.toast(role==='admin'?'관리자로 지정했습니다':'관리자 권한을 해제했습니다'); pAccounts(); };
  main.querySelectorAll('[data-gr]').forEach(b => b.onclick = () => { if(confirm('이 직원에게 관리자 권한을 줄까요?\n금액·정산을 포함한 모든 화면을 볼 수 있게 됩니다.')) setRole(b.dataset.gr,'admin'); });
  main.querySelectorAll('[data-rv]').forEach(b => b.onclick = () => { if(confirm('관리자 권한을 해제할까요?\n직원 계정으로 바뀝니다.')) setRole(b.dataset.rv,'staff'); });
  main.querySelectorAll('[data-ok]').forEach(b => b.onclick = () => setStatus(b.dataset.ok,'approved'));
  main.querySelectorAll('[data-no]').forEach(b => b.onclick = () => { if(confirm('이 계정을 사용할 수 없게 할까요?')) setStatus(b.dataset.no,'rejected'); });
  S.workers.forEach(w => { const g = $('#gd-'+w.id); if(g) g.onchange = async () => { const r = await sb.from('workers').update({gender:g.value||null}).eq('id',w.id); if(r.error) Y.toast(errMsg(r.error),1); else { w.gender = g.value||null; const b = g.closest('.ic').querySelector('.ich b'); b.className = 'gn ' + gcls(w.gender); Y.toast('저장했습니다'); } }; });
  S.workers.forEach(w => { const el = $('#tx-'+w.id); if(el) el.onchange = async () => { const r = await sb.from('workers').update({tax_type:el.value}).eq('id',w.id); if(r.error) Y.toast(errMsg(r.error),1); else { w.tax_type = el.value; Y.toast('저장했습니다'); } }; });
  main.querySelectorAll('[data-mm]').forEach(i => i.onchange = async () => { const r = await sb.from('workers').update({memo:i.value.trim()||null}).eq('id',i.dataset.mm); if(r.error) Y.toast(errMsg(r.error),1); else Y.toast('저장했습니다'); });
}

/* ---------- 라우팅 ---------- */
const PAGES = {board:pBoard,assign:pAssign,qr:pQr,notes:pNotes,query:pQuery,pay:pPay,items:pItems,accounts:pAccounts};
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
  $('#hb').onclick = tg; $('#ov').onclick = () => document.body.classList.remove('side-open');
  try{ if(innerWidth>=1000 && localStorage.getItem('yre_side')!=='0') document.body.classList.add('side-open'); }catch(e){}
  await loadRefs();
  window.addEventListener('hashchange', route);
  route();
})();
})();
