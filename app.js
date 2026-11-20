(function(){
  const C = window.YRE_CONFIG;
  const sb = window.supabase.createClient(C.SUPABASE_URL, C.SUPABASE_KEY, {auth:{persistSession:true,autoRefreshToken:true}});
  window.YRE = {
    sb, C,
    digits: s => String(s||'').replace(/\D/g,''),
    emailOf(phone){ return YRE.digits(phone) + '@' + C.EMAIL_DOMAIN; },
    fmtPhone(p){ p=YRE.digits(p); return p.length===11 ? p.replace(/(\d{3})(\d{4})(\d{4})/,'$1-$2-$3') : p; },
    validPhone(p){ p=YRE.digits(p); return /^01\d{8,9}$/.test(p); },
    async myProfile(){
      const {data:{session}} = await sb.auth.getSession();
      if(!session) return null;
      const {data,error} = await sb.from('profiles').select('id,name,phone,role,status').eq('id',session.user.id).maybeSingle();
      return error ? null : data;
    },
    async logout(){ await sb.auth.signOut(); location.reload(); }
  };
})();
