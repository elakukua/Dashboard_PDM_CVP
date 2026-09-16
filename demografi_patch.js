/* PATCH JS — Halaman "Demografi" (PDM CVP for GN Dashboard)
   -------------------------------------------------------------------------
   Semua blok di bawah ditempel di DALAM blok script pada index.html:

   [A] DEMO + merge   -> setelah  const FSP_ALL = [...];
   [B] page_demografi -> sebelum  const PAGES=[
   [C] ICO + PAGES    -> GANTI array PAGES yang lama
   [D] loop menu      -> GANTI  PAGES.forEach(([id,label])=>{ ... });

   Array DATA bawaan dashboard belum memuat kolom demografi (status sekolah,
   hubungan dengan KK, usia responden/KK, komposisi anggota RT, jenis
   kebutuhan khusus). Blok [A] menambahkannya dari raw export PDM_CVP.xlsx;
   urutan baris DEMO terverifikasi 1:1 dengan DATA untuk 46 record.
   ------------------------------------------------------------------------- */


/* ======================= [A] DATA DEMOGRAFI ============================= */
/* ---------- field demografi tambahan dari raw export PDM_CVP.xlsx ----------
   Urutan baris DEMO identik dengan urutan baris DATA (46 record, terverifikasi).
   Field: sekolah, usiaResp, hubKk, usiaKk, angg018, anggDws, anggKerja,
          anggHamil, disJenis[] */
const DEMO = [
  {"sekolah": "PAUD", "usiaResp": 45, "hubKk": "Isteri", "usiaKk": 44, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "TK", "usiaResp": 45, "hubKk": "Isteri", "usiaKk": 50, "angg018": 3, "anggDws": 4, "anggKerja": 3, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 47, "hubKk": null, "usiaKk": 47, "angg018": 1, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "TK", "usiaResp": 32, "hubKk": "Isteri", "usiaKk": 40, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "TK", "usiaResp": 45, "hubKk": "Isteri", "usiaKk": 43, "angg018": 1, "anggDws": 4, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 44, "hubKk": "Isteri", "usiaKk": 44, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "Putus Sekolah", "usiaResp": 60, "hubKk": null, "usiaKk": 60, "angg018": 1, "anggDws": 3, "anggKerja": 2, "anggHamil": 0, "disJenis": ["Berjalan"]},
  {"sekolah": "SD/sederajat", "usiaResp": 42, "hubKk": "Isteri", "usiaKk": 44, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMA/sederajat", "usiaResp": 50, "hubKk": "Isteri", "usiaKk": 51, "angg018": 1, "anggDws": 3, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 46, "hubKk": "Isteri", "usiaKk": 51, "angg018": 1, "anggDws": 3, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 56, "hubKk": null, "usiaKk": 56, "angg018": 1, "anggDws": 1, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 49, "hubKk": null, "usiaKk": 50, "angg018": 1, "anggDws": 2, "anggKerja": 3, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 45, "hubKk": "Isteri", "usiaKk": 46, "angg018": 2, "anggDws": 3, "anggKerja": 5, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 38, "hubKk": "Isteri", "usiaKk": 43, "angg018": 6, "anggDws": 2, "anggKerja": 0, "anggHamil": 0, "disJenis": []},
  {"sekolah": "PAUD", "usiaResp": 34, "hubKk": "Isteri", "usiaKk": 35, "angg018": 6, "anggDws": 2, "anggKerja": 0, "anggHamil": 0, "disJenis": []},
  {"sekolah": "PAUD", "usiaResp": 37, "hubKk": "Isteri", "usiaKk": 36, "angg018": 2, "anggDws": 2, "anggKerja": 4, "anggHamil": 1, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 45, "hubKk": null, "usiaKk": 45, "angg018": 3, "anggDws": 1, "anggKerja": 4, "anggHamil": 0, "disJenis": []},
  {"sekolah": "PAUD", "usiaResp": 26, "hubKk": "Isteri", "usiaKk": 38, "angg018": 3, "anggDws": 2, "anggKerja": 5, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 41, "hubKk": "Isteri", "usiaKk": 49, "angg018": 3, "anggDws": 2, "anggKerja": 5, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 40, "hubKk": null, "usiaKk": 40, "angg018": 4, "anggDws": 2, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "Putus Sekolah", "usiaResp": 40, "hubKk": null, "usiaKk": 50, "angg018": 3, "anggDws": 2, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 44, "hubKk": "Isteri", "usiaKk": 41, "angg018": 1, "anggDws": 2, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMA/sederajat", "usiaResp": 43, "hubKk": null, "usiaKk": 43, "angg018": 2, "anggDws": 2, "anggKerja": 3, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 35, "hubKk": null, "usiaKk": 35, "angg018": 2, "anggDws": 2, "anggKerja": 4, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 35, "hubKk": "Isteri", "usiaKk": 39, "angg018": 3, "anggDws": 2, "anggKerja": 5, "anggHamil": 0, "disJenis": []},
  {"sekolah": "Balita", "usiaResp": 35, "hubKk": "Isteri", "usiaKk": 39, "angg018": 3, "anggDws": 2, "anggKerja": 5, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 32, "hubKk": "Isteri", "usiaKk": 38, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 32, "hubKk": "Isteri", "usiaKk": 38, "angg018": 3, "anggDws": 2, "anggKerja": 0, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 36, "hubKk": null, "usiaKk": 36, "angg018": 2, "anggDws": 3, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 61, "hubKk": "Isteri", "usiaKk": 62, "angg018": 3, "anggDws": 5, "anggKerja": 8, "anggHamil": 0, "disJenis": []},
  {"sekolah": "TK", "usiaResp": 61, "hubKk": null, "usiaKk": 61, "angg018": 3, "anggDws": 3, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 26, "hubKk": "Isteri", "usiaKk": 31, "angg018": 3, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMA/sederajat", "usiaResp": 39, "hubKk": "Isteri", "usiaKk": 39, "angg018": 4, "anggDws": 3, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 39, "hubKk": null, "usiaKk": 39, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 46, "hubKk": null, "usiaKk": 46, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "Balita", "usiaResp": 26, "hubKk": null, "usiaKk": 26, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMA/sederajat", "usiaResp": 38, "hubKk": null, "usiaKk": 41, "angg018": 5, "anggDws": 0, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMA/sederajat", "usiaResp": 48, "hubKk": "Isteri", "usiaKk": 51, "angg018": 2, "anggDws": 2, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 34, "hubKk": "Isteri", "usiaKk": 38, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 54, "hubKk": null, "usiaKk": 54, "angg018": 1, "anggDws": 2, "anggKerja": 2, "anggHamil": 0, "disJenis": ["Berkomunikasi"]},
  {"sekolah": "SD/sederajat", "usiaResp": 46, "hubKk": "Isteri", "usiaKk": 48, "angg018": 2, "anggDws": 3, "anggKerja": 2, "anggHamil": 0, "disJenis": []},
  {"sekolah": "PAUD", "usiaResp": 33, "hubKk": "Isteri", "usiaKk": 39, "angg018": 3, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": ["Mengurus diri"]},
  {"sekolah": "SD/sederajat", "usiaResp": 25, "hubKk": "Isteri", "usiaKk": 31, "angg018": 3, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 33, "hubKk": "Isteri", "usiaKk": 37, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SD/sederajat", "usiaResp": 36, "hubKk": "Isteri", "usiaKk": 35, "angg018": 2, "anggDws": 2, "anggKerja": 1, "anggHamil": 0, "disJenis": []},
  {"sekolah": "SMP/sederajat", "usiaResp": 37, "hubKk": "Isteri", "usiaKk": 38, "angg018": 4, "anggDws": 3, "anggKerja": 2, "anggHamil": 0, "disJenis": []}
];
DEMO.forEach((d,i)=>{ if(DATA[i]) Object.assign(DATA[i],d); });


/* ============== [B] KONSTANTA, HELPER & page_demografi ================= */
/* ================= AREA 0 — DEMOGRAFI ================= */
/* ---------- konstanta urutan kategori (mengikuti XLSForm) ---------- */
const SEKOLAH=[['Balita','Balita / belum sekolah'],['PAUD','PAUD'],['TK','TK'],
  ['SD/sederajat','SD / sederajat'],['SMP/sederajat','SMP / sederajat'],
  ['SMA/sederajat','SMA / sederajat'],['Putus Sekolah','Putus sekolah']];
const PENDKK=['Tidak/belum sekolah','SD','SMP','SMA','Perguruan Tinggi'];
const RUMAHST=['Milik sendiri','Kontrak','Menumpang','Lainnya'];
const DISJEN=['Melihat','Mendengar','Berjalan','Mengingat','Mengurus diri','Berkomunikasi','Lainnya'];
const UKEL=[['0\u20135 th',0,5],['6\u201311 th',6,11],['12\u201314 th',12,14],['15\u201318 th',15,18]];

/* ---------- helper tambahan ---------- */
const avg = a => a.length ? a.reduce((x,y)=>x+y,0)/a.length : null;
const rng = a => a.length ? Math.min(...a)+'\u2013'+Math.max(...a) : '\u2014';
/* blok angka ringkas: [{v,t}] */
function stat(items){
  return '<div class="dgs">'+items.map(i=>
    '<div><b>'+i.v+'</b><small>'+i.t+'</small></div>').join('')+'</div>';
}
/* distribusi nilai numerik -> bar (mis. ukuran RT, jumlah anak) */
function numBars(vals,unit,base){
  return [...new Set(vals)].sort((x,y)=>x-y).map(k=>{
    const v=vals.filter(x=>x===k).length;
    return {k:k+' '+unit,v:v,lbl:of_(v,base)};
  });
}

/* ================= HALAMAN DEMOGRAFI ================= */
function page_demografi(R){
  const n=R.length;

  /* --- anak penerima --- */
  const aP=C(R,r=>r.jkAnak==='Perempuan'), aL=C(R,r=>r.jkAnak==='Laki-laki');
  const ua=R.map(r=>r.usiaAnak).filter(v=>v!=null);
  const uaBars=UKEL.map(([l,lo,hi])=>{
    const v=C(R,r=>r.usiaAnak!=null && r.usiaAnak>=lo && r.usiaAnak<=hi);
    return {k:l,v:v,lbl:of_(v,n)};
  });
  const skBars=SEKOLAH.map(([k,l])=>{
    const v=C(R,r=>r.sekolah===k);
    return {k:l,v:v,cls:(k==='Putus Sekolah'?'dk':''),lbl:of_(v,n)};
  });
  const sekNA=C(R,r=>!r.sekolah);

  /* --- responden & kepala keluarga --- */
  const rP=C(R,r=>r.jkResp==='Perempuan'), rL=C(R,r=>r.jkResp==='Laki-laki');
  const kP=C(R,r=>r.jkKk==='Perempuan'),  kL=C(R,r=>r.jkKk==='Laki-laki');
  const ur=R.map(r=>r.usiaResp).filter(v=>v!=null);
  const uk=R.map(r=>r.usiaKk).filter(v=>v!=null);
  const jkBars=[
    {k:'Responden \u2014 Perempuan',v:rP,lbl:of_(rP,n)},
    {k:'Responden \u2014 Laki-laki',v:rL,cls:'l',lbl:of_(rL,n)},
    {k:'Kepala keluarga \u2014 Perempuan',v:kP,cls:'dk',lbl:of_(kP,n)},
    {k:'Kepala keluarga \u2014 Laki-laki',v:kL,cls:'ll',lbl:of_(kL,n)}
  ];

  /* --- hubungan responden dengan KK (pengambil keputusan rumah tangga) --- */
  const hSelf=C(R,r=>r.isKk==='Ya');
  const hMap={};
  R.filter(r=>r.isKk!=='Ya').forEach(r=>{
    const k=r.hubKk || 'Tidak diisi (skip logic)';
    hMap[k]=(hMap[k]||0)+1;
  });
  const hubBars=[{k:'Kepala keluarga (responden sendiri)',v:hSelf,cls:'dk',lbl:of_(hSelf,n)}]
    .concat(Object.keys(hMap).sort((a,b)=>hMap[b]-hMap[a]).map(k=>({
      k:esc(k),v:hMap[k],
      cls:(k.indexOf('Tidak diisi')===0?'g':''),
      lbl:of_(hMap[k],n)})));
  const hubNA=Object.keys(hMap).filter(k=>k.indexOf('Tidak diisi')===0)
    .reduce((a,k)=>a+hMap[k],0);

  /* --- pendidikan KK --- */
  const pdBars=PENDKK.map(k=>{
    const v=C(R,r=>r.pendKk===k);
    return {k:k,v:v,cls:(k==='Tidak/belum sekolah'?'dk':''),lbl:of_(v,n)};
  });
  const pdLow=C(R,r=>r.pendKk==='Tidak/belum sekolah'||r.pendKk==='SD');

  /* --- rumah tangga --- */
  const tot=R.map(r=>r.anggota).filter(v=>v!=null);
  const a18=R.map(r=>r.angg018).filter(v=>v!=null);
  const adw=R.map(r=>r.anggDws).filter(v=>v!=null);
  const hamil=C(R,r=>r.anggHamil>0);
  /* DQ: jumlah anggota berpenghasilan > jumlah anggota dewasa (>18) \u2192 tidak plausibel */
  const keOk=R.filter(r=>r.anggKerja!=null && r.anggDws!=null && r.anggKerja<=r.anggDws)
              .map(r=>r.anggKerja);
  const keDq=C(R,r=>r.anggKerja!=null && r.anggDws!=null && r.anggKerja>r.anggDws);
  const rmBars=RUMAHST.map(k=>{
    const v=C(R,r=>r.rumah===k);
    return {k:k,v:v,cls:(k==='Lainnya'?'l':(k==='Menumpang'?'dk':'')),lbl:of_(v,n)};
  });

  /* --- inklusi & disabilitas --- */
  const dRow=R.filter(r=>r.disab==='Ya'), dYa=dRow.length;
  const dTdk=C(R,r=>r.disab==='Tidak');
  /* penyebut = RT disabilitas (sangat kecil) \u2192 label angka absolut, bukan persen */
  const dBars=DISJEN.map(k=>{
    const v=dRow.filter(r=>(r.disJenis||[]).indexOf(k)>=0).length;
    return {k:k,v:v,lbl:String(v)};
  });

  /* ---------- KPI ---------- */
  const K=kpi([
    {t:'Responden (RC)',v:n,b:'RC terwawancara pada seleksi ini'},
    {t:'Median usia anak',v:(ua.length?dec(med(ua),1):'\u2014')+'<span class="u"> th</span>',
      b:'rentang '+rng(ua)+' th \u00b7 basis '+ua.length+' RC'},
    {t:'Anak perempuan',v:aP,b:pctS(aP,n)+' dari '+n+' RC'},
    {t:'Responden perempuan',v:rP,b:pctS(rP,n)+' dari '+n+' RC'},
    {t:'RT dengan anggota disabilitas',v:dYa,b:pctS(dYa,n)+' dari '+n+' RC'}
  ]);

  /* ---------- 10 kartu demografi, 2 kolom ---------- */
  const G=[
    card(null,'Jenis kelamin anak penerima','Basis '+n+' anak penerima GN',
      donut(aP,aL,
        'Perempuan <em>('+pctS(aP,n)+')</em>',
        'Laki-laki <em>('+pctS(aL,n)+')</em>',
        'Anak perempuan'),null),

    card(null,'Kelompok usia anak penerima',
      'Median '+(ua.length?dec(med(ua),1):'\u2014')+' th \u00b7 rentang '+rng(ua)+' th',
      hb(uaBars,{max:n}),null),

    card(null,'Status sekolah anak saat ini',
      'Urutan jenjang'+(sekNA?' \u00b7 <span class="dq">'+sekNA+' tanpa jawaban</span>':''),
      hb(skBars,{max:n,tight:true}),null),

    card(null,'Jumlah anak usia 0\u201318 tahun dalam RT',
      'Rata-rata '+(a18.length?dec(avg(a18),1):'\u2014')+' anak per RT',
      hb(numBars(a18,'anak',n),{max:n,tight:true}),null),

    card(null,'Hubungan responden dengan kepala keluarga',
      'Pengambil keputusan rumah tangga'+
      (hubNA?' \u00b7 <span class="dq">'+hubNA+' bukan KK tapi hubungan kosong</span>':''),
      hb(hubBars,{max:n}),null),

    card(null,'Jenis kelamin responden &amp; kepala keluarga',
      'Median usia responden '+(ur.length?dec(med(ur),0):'\u2014')+' th \u00b7 median usia KK '+
      (uk.length?dec(med(uk),0):'\u2014')+' th',
      hb(jkBars,{max:n}),null),

    card(null,'Pendidikan terakhir kepala keluarga',
      pdLow+' dari '+n+' RC \u2014 KK berpendidikan SD atau lebih rendah ('+pctS(pdLow,n)+')',
      hb(pdBars,{max:n}),null),

    card(null,'Ukuran rumah tangga',
      'Rata-rata '+(tot.length?dec(avg(tot),1):'\u2014')+' anggota per RT',
      stat([
        {v:(tot.length?dec(avg(tot),1):'\u2014'),t:'Rata-rata anggota RT'},
        {v:(a18.length?dec(avg(a18),1):'\u2014'),t:'Rata-rata anggota usia 0\u201318 th'},
        {v:(adw.length?dec(avg(adw),1):'\u2014'),t:'Rata-rata anggota usia >18 th'}
      ])+hb(numBars(tot,'orang',n),{max:n,tight:true}),null),

    card(null,'Status tempat tinggal',
      'Kepemilikan rumah yang ditempati RC',
      hb(rmBars,{max:n}),null),

    card(null,'Inklusi &amp; disabilitas',
      'Jenis kebutuhan khusus memakai angka absolut \u2014 penyebut hanya '+dYa+' RT',
      stat([
        {v:dYa,t:'RT dengan anggota disabilitas'},
        {v:pctS(dYa,n),t:'Proporsi dari '+n+' RC'},
        {v:dTdk,t:'RT tanpa anggota disabilitas'}
      ])+hb(dBars,{max:Math.max(1,dYa),tight:true}),null)
  ];

  /* ---------- tabel demografi per AP (untuk tindak lanjut lapangan) ---------- */
  const apBody=byAp(R).map(([a,rr])=>{
    const nn=rr.length;
    const p=C(rr,r=>r.jkAnak==='Perempuan');
    const u=rr.map(r=>r.usiaAnak).filter(v=>v!=null);
    const g=rr.map(r=>r.anggota).filter(v=>v!=null);
    return '<tr><td>'+esc(a)+'</td>'+
      '<td class="n">'+nn+'</td>'+
      '<td class="n">'+p+'</td>'+
      '<td class="n">'+(nn-p)+'</td>'+
      '<td class="n">'+(u.length?dec(med(u),1):'\u2014')+'</td>'+
      '<td class="n">'+C(rr,r=>r.pendKk==='Tidak/belum sekolah'||r.pendKk==='SD')+'</td>'+
      '<td class="n">'+C(rr,r=>r.rumah==='Milik sendiri')+'</td>'+
      '<td class="n">'+C(rr,r=>r.disab==='Ya')+'</td>'+
      '<td class="n">'+(g.length?dec(avg(g),1):'\u2014')+'</td></tr>';
  }).join('');
  const apCard=card('Tabel A','Profil demografi per Area Program',
    'Angka absolut (n per AP kecil) \u00b7 median usia anak dan rata-rata anggota RT dalam desimal',
    tbl('<tr><th>Area Program</th><th class="n">RC</th><th class="n">Anak<br>P</th>'+
        '<th class="n">Anak<br>L</th><th class="n">Median usia<br>anak (th)</th>'+
        '<th class="n">KK \u2264 SD</th><th class="n">Rumah milik<br>sendiri</th>'+
        '<th class="n">RT<br>disabilitas</th><th class="n">Rata-rata<br>anggota RT</th></tr>',
        apBody),null);

  /* catatan DQ khusus anggota berpenghasilan \u2014 dipisah agar tidak mengotori kartu */
  const dqBox = keDq ?
    '<div class="note"><b>Catatan kualitas data.</b> Kolom <i>jumlah anggota keluarga yang '+
    'berpenghasilan</i> tidak dipakai sebagai indikator pada halaman ini: <b>'+keDq+' dari '+n+
    ' record</b> mencatat jumlah berpenghasilan melebihi jumlah anggota usia &gt;18 tahun '+
    '(pola pengisian = total anggota RT). Rata-rata pada record yang lolos uji: '+
    (keOk.length?dec(avg(keOk),1):'\u2014')+' orang (basis '+keOk.length+' RC). '+
    'Perlu verifikasi lapangan sebelum dilaporkan.</div>' : '';

  return K + '<div class="cv-grid">'+G.join('')+'</div>' + row('r-1',[apCard]) + dqBox;
}


/* ============ [C] REGISTRASI HALAMAN + IKON NAVIGASI ================== */
/* ikon menu: inline SVG. Atribut width/height/fill/stroke ditulis langsung pada
   elemen svg agar ikon tetap benar walau CSS belum termuat. */
const ICO={
  cover:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/></svg>',
  grid :'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>',
  users:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M16 19v-1a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v1"/><circle cx="9" cy="7" r="3.2"/><path d="M22 19v-1a4 4 0 0 0-3-3.87"/><path d="M16.5 4.2a3.2 3.2 0 0 1 0 5.6"/></svg>',
  info :'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><path d="M12 8h.01"/></svg>',
  bank :'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 10 12 4l9 6"/><path d="M5 10v9"/><path d="M19 10v9"/><path d="M3 20h18"/><path d="M10 20v-6h4v6"/></svg>',
  truck:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="7" width="12" height="9" rx="1"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6.5" cy="18" r="1.8"/><circle cx="17" cy="18" r="1.8"/></svg>',
  wallet:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10.5h18"/><path d="M16.5 14.8h.01"/></svg>',
  smile:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M8.5 14.2a4.4 4.4 0 0 0 7 0"/><path d="M9 9.5h.01"/><path d="M15 9.5h.01"/></svg>',
  shield:'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l8 3v6c0 4.5-3.2 7.6-8 9-4.8-1.4-8-4.5-8-9V6z"/><path d="M8.8 12.2l2.3 2.3 4.1-4.6"/></svg>'
};

/* [id, label menu, H1, sub-judul, fungsi render, ikon] */
const PAGES=[
 ['p0','Cover',null,null,cover,ICO.cover],
 ['p1','Overview','Overview','Ringkasan 6 area asesmen',pg1,ICO.grid],
 ['pd','Demografi','Demografi Responden &amp; Rumah Tangga','Profil anak penerima, responden, kepala keluarga, dan rumah tangga',page_demografi,ICO.users],
 ['p2','Informasi','Area 1 \u2014 Penyediaan Informasi','Tabel 1 &amp; 2',pg2,ICO.info],
 ['p3','Kinerja FSP','Area 2 \u2014 Kinerja Penyedia Jasa Keuangan (FSP)','Tabel 3 &amp; 4',pg3,ICO.bank],
 ['p4','Distribusi','Area 3 \u2014 Proses Distribusi','Tabel 5, 6 &amp; 7',pg4,ICO.truck],
 ['p5','Penggunaan','Area 4 \u2014 Penggunaan Dana','Tabel 8',pg5,ICO.wallet],
 ['p6','Kepuasan','Area 5 \u2014 Kepuasan &amp; Perbandingan Modalitas','Tabel 9',pg6,ICO.smile],
 ['p7','Akuntabilitas','Area 6 \u2014 Akuntabilitas','Tabel 11',pg7,ICO.shield]
];


/* ==================== [D] LOOP PEMBANGUN MENU ========================= */
PAGES.forEach(([id,label,,,,icon])=>{
  const b=document.createElement('button');
  b.innerHTML=(icon||'')+'<span>'+label+'</span>';
  b.dataset.p=id;
  b.setAttribute('aria-current',String(id===COVER_ID));
  b.addEventListener('click',()=>go(id));
  $('nav').appendChild(b);
});
