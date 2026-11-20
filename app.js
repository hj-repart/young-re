(function(){
  const C = window.YRE_CONFIG;
  const sb = window.supabase.createClient(C.SUPABASE_URL, C.SUPABASE_KEY, {auth:{persistSession:true,autoRefreshToken:true}});
  const DOW = ['일','월','화','수','목','금','토'];
  const Y = window.YRE = {
    sb, C, lastError:'',
    digits: s => String(s||'').replace(/\D/g,''),
    emailOf(phone){ return Y.digits(phone) + '@' + C.EMAIL_DOMAIN; },
    fmtPhone(p){ p=Y.digits(p); return p.length===11 ? p.replace(/(\d{3})(\d{4})(\d{4})/,'$1-$2-$3') : (p.length===10 ? p.replace(/(\d{3})(\d{3})(\d{4})/,'$1-$2-$3') : p); },
    validPhone(p){ p=Y.digits(p); return /^01\d{8,9}$/.test(p); },
    esc: s => String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])),
    today(){ return new Date(Date.now()+9*36e5).toISOString().slice(0,10); },
    addDays(d,n){ const t=new Date(d+'T00:00:00Z'); t.setUTCDate(t.getUTCDate()+n); return t.toISOString().slice(0,10); },
    dow(d){ return DOW[new Date(d+'T00:00:00Z').getUTCDay()]; },
    fmtDate(d){ return String(d).slice(0,10).replace(/-/g,'.') + ' (' + Y.dow(String(d).slice(0,10)) + ')'; },
    fmtTime(iso){ return iso ? new Date(iso).toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit',hour12:false,timeZone:'Asia/Seoul'}) : '-'; },
    fmtDur(a,b){ if(!a||!b) return '-'; const m=Math.round((new Date(b)-new Date(a))/60000); if(m<0) return '-'; return Math.floor(m/60)+'시간 '+String(m%60).padStart(2,'0')+'분'; },
    won(n){ return Number(n||0).toLocaleString('ko-KR')+'원'; },
    toLocalInput(iso){ return iso ? new Date(new Date(iso).getTime()+9*36e5).toISOString().slice(0,16) : ''; },
    fromLocalInput(s){ return s ? new Date(s+':00+09:00').toISOString() : null; },
    toast(msg,bad){
      let t=document.getElementById('toast'); if(!t){ t=document.createElement('div'); t.id='toast'; document.body.append(t); }
      t.textContent=msg; t.className='toast show'+(bad?' bad':'');
      clearTimeout(Y._tt); Y._tt=setTimeout(()=>{ t.className='toast'; }, bad?4500:2200);
    },
    async myProfile(){
      const {data:{session}} = await sb.auth.getSession();
      if(!session) return null;
      const {data,error} = await sb.from('profiles').select('id,name,phone,role,status').eq('id',session.user.id).maybeSingle();
      Y.lastError = error ? (error.message || String(error)) : (data ? '' : '계정 정보(profiles)가 없습니다');
      return error ? null : data;
    },
    async guard(role){
      const p = await Y.myProfile();
      if(!p || p.status!=='approved' || (role==='admin' && p.role!=='admin')){ location.replace('index.html'); return null; }
      return p;
    },
    async logout(){ await sb.auth.signOut(); location.href='index.html'; }
  };
  if('serviceWorker' in navigator){ window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(()=>{})); }
})();
