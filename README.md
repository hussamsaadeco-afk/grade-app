<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>نظام إدارة المدرسة الاحترافي</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root { --bg: #f1f5f9; --card: #ffffff; --text: #1e293b; --p: #3b82f6; --d: #1e3a8a; --s: #10b981; --e: #ef4444; --w: #f59e0b; --b: #e2e8f0; }
[data-theme="dark"] { --bg: #0f172a; --card: #1e293b; --text: #f8fafc; --b: #334155; }
* { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, sans-serif; }
body { background: var(--bg); color: var(--text); padding-bottom: 80px; }
#login, #st-portal, #modal { position: fixed; inset: 0; background: var(--d); display: flex; align-items: center; justify-content: center; z-index: 9999; padding: 20px; }
#modal { background: rgba(0,0,0,0.7); display: none; z-index: 10000; }
.box { background: var(--card); padding: 25px; border-radius: 16px; width: 100%; max-width: 400px; text-align: center; max-height: 90vh; overflow-y: auto; }
input, select, button { width: 100%; padding: 10px; margin-bottom: 10px; border: 1px solid var(--b); border-radius: 8px; font-size: 14px; background: var(--card); color: var(--text); outline: none; }
button { cursor: pointer; font-weight: bold; border: none; background: var(--p); color: #fff; }
button.o { background: transparent; color: var(--p); border: 2px solid var(--p); }
button.s { background: var(--s); } button.w { background: var(--w); } button.d { background: var(--e); }
header { background: var(--d); color: #fff; padding: 15px; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 100; }
header h1 { font-size: 16px; }
.h-btns { display: flex; gap: 8px; }
main { padding: 10px; max-width: 1000px; margin: 0 auto; }
.tab { display: none; } .tab.active { display: block; }
.card { background: var(--card); border-radius: 12px; padding: 15px; margin-bottom: 10px; border: 1px solid var(--b); }
.card h2 { font-size: 15px; color: var(--p); margin-bottom: 10px; border-bottom: 2px solid var(--b); padding-bottom: 8px; }
.row { display: flex; gap: 10px; flex-wrap: wrap; } .row > * { flex: 1; min-width: 120px; }
.li { background: var(--card); border: 1px solid var(--b); padding: 10px; margin-bottom: 8px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }
.li b { color: var(--p); } .li small { opacity: 0.7; font-size: 11px; display: block; margin-top: 2px; }
.ib { width: auto; padding: 5px 10px; margin: 0; border-radius: 6px; font-size: 12px; }
.nav { position: fixed; bottom: 0; left: 0; right: 0; background: var(--card); display: flex; justify-content: space-around; padding: 8px 0; border-top: 1px solid var(--b); z-index: 100; overflow-x: auto; }
.nb { background: none; display: flex; flex-direction: column; align-items: center; color: var(--text); opacity: 0.6; font-size: 16px; width: auto; padding: 5px; margin: 0; white-space: nowrap; }
.nb span { font-size: 10px; margin-top: 3px; font-weight: bold; } .nb.active { color: var(--p); opacity: 1; }
table { width: 100%; border-collapse: collapse; font-size: 12px; background: var(--card); }
th, td { border: 1px solid var(--b); padding: 6px; text-align: center; }
th { background: var(--bg); color: var(--p); }
#gt table { display: block; overflow-x: auto; white-space: nowrap; }
#gt th:first-child, #gt td:first-child { position: sticky; right: 0; background: var(--card); border-left: 2px solid var(--b); }
.dash { background: var(--bg); border: 1px solid var(--b); border-radius: 12px; padding: 15px; text-align: center; flex: 1; }
.dash h3 { font-size: 24px; color: var(--p); }
.rc { background: var(--card); border-radius: 12px; padding: 15px; border: 1px solid var(--b); margin-top: 10px; }
.rh { border-bottom: 2px dashed var(--b); padding-bottom: 10px; margin-bottom: 10px; text-align: center; }
.si { display: flex; justify-content: space-around; flex-wrap: wrap; gap: 5px; font-size: 12px; font-weight: bold; }
.gg { display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 10px; margin-top: 10px; }
.sc { border: 1px solid var(--b); border-radius: 8px; padding: 10px; background: var(--bg); }
.st { font-weight: bold; color: var(--p); border-bottom: 1px solid var(--b); padding-bottom: 5px; margin-bottom: 5px; text-align: center; }
.gr { display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 2px; }
.gr span { font-weight: bold; color: var(--p); }
.fb { background: var(--s); color: #fff; text-align: center; padding: 6px; border-radius: 6px; font-weight: bold; margin-top: 6px; }
.fb.f { background: var(--e); }
.footer { text-align: center; padding: 20px; margin-top: 20px; color: #64748b; font-size: 12px; border-top: 1px solid var(--b); }
.footer strong { color: var(--p); }
.toast { position: fixed; bottom: 80px; left: 50%; transform: translateX(-50%) translateY(100px); background: var(--d); color: #fff; padding: 10px 20px; border-radius: 25px; font-size: 13px; opacity: 0; transition: all .3s; z-index: 10001; }
.toast.show { transform: translateX(-50%) translateY(0); opacity: 1; }
</style>
</head>
<body>

<div id="login">
  <div class="box">
    <h2>🔐 دخول الإدارة</h2>
    <input type="text" id="lp" placeholder="كلمة المرور (admin)">
    <div id="lerr" style="color:var(--e);font-size:12px;display:none;margin-bottom:10px;">كلمة المرور غير صحيحة!</div>
    <button onclick="loginAdmin()">دخول</button>
    <button class="o" onclick="loginAdmin(true)" style="font-size:12px;">(دخول تجريبي)</button>
    <button class="o" onclick="showSP()">🎓 دخول الطالب</button>
  </div>
</div>

<div id="st-portal" style="display:none;">
  <div class="box">
    <h2>🎓 بوابة الطالب</h2>
    <input type="text" id="sc" placeholder="كود الطالب">
    <button onclick="loginSt()">عرض</button>
    <button class="o" onclick="hideSP()">🔙 رجوع</button>
  </div>
</div>

<div id="st-view" style="display:none;padding:15px;max-width:1000px;margin:0 auto;">
  <div id="st-name" style="font-size:20px;font-weight:bold;color:var(--p);text-align:center;margin-bottom:15px;"></div>
  <div id="st-grades"></div>
  <button class="d" onclick="document.getElementById('st-view').style.display='none';showSP();" style="margin-top:15px;">خروج</button>
</div>

<div id="modal">
  <div class="box" style="max-width: 600px;">
    <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
      <h3 style="color:var(--p);">ملف الطالب الشامل</h3>
      <button class="ib d" onclick="closeModal()" style="width:auto;">X</button>
    </div>
    <div id="modal-content"></div>
  </div>
</div>

<div id="admin" style="display:none;">
  <header>
    <h1>🏫 إدارة المدرسة</h1>
    <div class="h-btns">
      <button class="ib" style="background:rgba(255,255,255,.2);" onclick="location.reload()">🔄</button>
      <button class="ib" style="background:rgba(255,255,255,.2);" onclick="toggleTheme()">🌓</button>
      <button class="ib" style="background:rgba(255,255,255,.2);" onclick="logout()">خروج</button>
    </div>
  </header>
  <nav class="nav">
    <button class="nb active" data-tab="dash">📊<span>الرئيسية</span></button>
    <button class="nb" data-tab="students">👨‍🎓<span>الطلاب</span></button>
    <button class="nb" data-tab="grades">📝<span>الدرجات</span></button>
    <button class="nb" data-tab="att">📅<span>الغياب</span></button>
    <button class="nb" data-tab="act">🏅<span>النشاط</span></button>
    <button class="nb" data-tab="rep">📄<span>الكشوف</span></button>
    <button class="nb" data-tab="levels">📈<span>المستوى</span></button>
    <button class="nb" data-tab="set">⚙️<span>الإدارة</span></button>
  </nav>
  <main>
    <section id="dash" class="tab active">
      <div class="card">
        <h2>📊 لوحة القيادة</h2>
        <div class="row" style="margin-bottom:10px;"><div class="dash"><h3 id="ds">0</h3><p>الطلاب</p></div><div class="dash"><h3 id="da">0</h3><p>غياب اليوم</p></div></div>
        <button class="s" onclick="showStats()" style="margin-bottom:10px;">📊 إحصائيات الصفوف والجنسين</button>
        <div id="stats-box"></div>
      </div>
    </section>
    <section id="students" class="tab">
      <div class="card">
        <h2>👨‍🎓 الطلاب</h2>
        <div class="row"><select id="stl" onchange="uSF()"><option value="">-- الصف --</option></select><select id="sts" onchange="rS()"><option value="">-- الشعبة --</option></select></div>
        <input id="srch" placeholder="🔍 بحث..." oninput="rS()">
        <div id="stl-list"></div>
        <button class="s" style="margin-top:10px;" onclick="printL()">🖨️ طباعة الأسماء</button>
        <div class="card" style="margin-top:10px;background:#dcfce7;">
          <h3 style="font-size:13px;color:#065f46;margin-bottom:5px;">🎓 الترحيل الجماعي</h3>
          <div class="row"><select id="ufl"><option value="">-- من صف --</option></select><select id="utl"><option value="">-- إلى صف --</option></select></div>
          <button class="s" onclick="upgradeAll()">ترقية الناجحين</button>
        </div>
        <div class="card" style="margin-top:10px;background:#fef08a;">
          <h3 style="font-size:13px;color:#854d0e;margin-bottom:5px;">🚚 نقل الطلاب</h3>
          <div class="row"><select id="mvl" onchange="uMS()"><option value="">-- الصف --</option></select><select id="mvs"><option value="">-- الشعبة --</option></select></div>
          <button class="w" onclick="moveSel()">نقل المحددين</button>
        </div>
      </div>
    </section>
    <section id="grades" class="tab">
      <div class="card">
        <h2>📝 الدرجات</h2>
        <div class="row"><select id="gl" onchange="uGS()"><option value="">-- الصف --</option></select><select id="gs" onchange="rG()"><option value="">-- الشعبة --</option></select></div>
        <select id="gsub" onchange="rG()"><option value="">-- المادة --</option></select>
        <button class="s" onclick="importG()">📥 استيراد من إكسل</button>
        <input type="file" id="gfile" accept=".xlsx,.xls,.csv" style="display:none;" onchange="procG(this)">
        <div id="gt"></div>
      </div>
    </section>
    <section id="att" class="tab">
      <div class="card">
        <h2>📅 الغياب</h2>
        <input type="date" id="ad" onchange="rA()">
        <div class="row"><select id="al" onchange="uAS()"><option value="">-- الصف --</option></select><select id="as" onchange="rA()"><option value="">-- الشعبة --</option></select></div>
        <input id="sa" placeholder="🔍 بحث..." oninput="rA()">
        <div id="att-list"></div>
      </div>
    </section>
    <section id="act" class="tab">
      <div class="card">
        <h2>🏅 النشاط</h2>
        <div class="row"><select id="acl" onchange="uACS()"><option value="">-- الصف --</option></select><select id="acs" onchange="rAC()"><option value="">-- الشعبة --</option></select></div>
        <input id="sac" placeholder="🔍 بحث..." oninput="rAC()">
        <div id="act-list"></div>
        <button class="s" id="sact" style="display:none;margin-top:10px;" onclick="saveAC()">💾 حفظ</button>
      </div>
    </section>
    <section id="rep" class="tab">
      <div class="card">
        <h2>📄 الكشوفات</h2>
        <div class="row"><select id="rl" onchange="uRS()"><option value="">-- الصف --</option></select><select id="rs"><option value="">-- الشعبة --</option></select></div>
        <select id="rtype" style="margin-bottom:10px;">
          <option value="grades">كشف درجات الشعبة</option>
          <option value="att">كشف غياب الشعبة</option>
          <option value="act">كشف نشاط الشعبة</option>
        </select>
        <button class="p" onclick="genRep()" style="background:var(--p);">عرض الكشف</button>
        <div id="rbox"></div>
      </div>
    </section>
    <section id="levels" class="tab">
      <div class="card">
        <h2>📈 المستوى العلمي</h2>
        <div class="row"><select id="ll" onchange="uLS()"><option value="">-- الصف --</option></select><select id="ls" onchange="rL()"><option value="">-- الشعبة --</option></select></div>
        <select id="lf" onchange="rL()" style="margin-bottom:10px;"><option value="all">المعدل العام</option><option value="m1">الشهر الأول</option><option value="m2">الشهر الثاني</option><option value="mid">نصف السنة</option><option value="final">النهائي</option></select>
        <div id="lvl-list"></div>
      </div>
    </section>
    <section id="set" class="tab">
      <div class="card">
        <h2>🏫 الصفوف والشعب</h2>
        <div class="row"><input id="nl" placeholder="اسم الصف"><button class="s" onclick="addL()" style="width:80px;margin:0;">إضافة</button></div>
        <div id="llist"></div>
        <div class="row" style="margin-top:10px;"><select id="sl" onchange="uSS()"><option value="">-- اختر الصف --</option></select><input id="ns" placeholder="الشعبة"><button class="s" onclick="addS()" style="width:80px;margin:0;">إضافة</button></div>
        <div id="slist"></div>
      </div>
      <div class="card">
        <h2>📥 استيراد الطلاب</h2>
        <p style="font-size:12px;margin-bottom:5px;">(الاسم، الشعبة، الجنس، رقم ولي الأمر)</p>
        <div class="row"><select id="il" onchange="uIS()"><option value="">-- الصف --</option></select><select id="is"><option value="">-- الشعبة --</option></select></div>
        <select id="ig" style="margin-bottom:10px;"><option value="ذكر">ذكور</option><option value="أنثى">إناث</option><option value="auto">من الإكسل</option></select>
        <input type="file" id="ef" accept=".xlsx,.xls,.csv">
        <button class="s" onclick="importS()">استيراد</button>
      </div>
      <div class="card"><h2>📚 المواد</h2><div class="row"><input id="sn" placeholder="اسم المادة"><button class="s" onclick="addSub()" style="width:80px;margin:0;">إضافة</button></div><div id="sublist"></div></div>
      <div class="card"><h2>💾 تصدير</h2><button class="w" onclick="expCSV()">📤 تصدير Excel</button></div>
      <div class="card"><h2>⚠️ الخطر</h2><button class="d" onclick="clearAll()">🗑️ حذف الكل</button></div>
    </section>
  </main>
</div>
<div id="toast" class="toast"></div>
<div class="footer">تم التصميم بواسطة <strong>د. حسام الخالدي</strong> &copy; 2027</div>
<script>
let DB;
try { let s = localStorage.getItem('schoolDB_v100'); if(s){ DB = JSON.parse(decodeURIComponent(atob(s))); } } catch(e){}
if(!DB) DB = {"students":[],"levels":[],"sections":[],"subjects":[{"id":1,"name":"اسلامية"},{"id":2,"name":"عربي"},{"id":3,"name":"انكليزي"},{"id":4,"name":"رياضيات"},{"id":5,"name":"احياء"},{"id":6,"name":"فيزياء"},{"id":7,"name":"كيمياء"},{"id":8,"name":"اجتماعيات"},{"id":9,"name":"اخلاقيه"}],"grades":[],"attendance":{},"activity":{},"adminPass":"admin"};
function sv(){ localStorage.setItem('schoolDB_v100', btoa(encodeURIComponent(JSON.stringify(DB)))); }
function tst(m,t='s'){ let e=document.getElementById('toast'); e.textContent=m; e.className='toast show '+t; setTimeout(()=>e.className='toast '+t,2500); }
function nid(a){ return a.length?Math.max(...a.map(x=>x.id))+1:1; }
function toggleTheme(){ let b=document.body; if(b.getAttribute('data-theme')==='dark'){ b.removeAttribute('data-theme'); localStorage.setItem('theme','light'); } else { b.setAttribute('data-theme','dark'); localStorage.setItem('theme','dark'); } }
if(localStorage.getItem('theme')==='dark') document.body.setAttribute('data-theme','dark');

let isL = localStorage.getItem('isLogged_v100')==='true';
if(isL){ document.getElementById('login').style.display='none'; document.getElementById('admin').style.display='block'; setTimeout(init,100); }
function loginAdmin(bypass=false){ let p=document.getElementById('lp').value; if(bypass || p===DB.adminPass || p==='admin'){ localStorage.setItem('isLogged_v100','true'); document.getElementById('login').style.display='none'; document.getElementById('admin').style.display='block'; init(); } else { document.getElementById('lerr').style.display='block'; } }
function logout(){ localStorage.setItem('isLogged_v100','false'); document.getElementById('admin').style.display='none'; document.getElementById('login').style.display='flex'; }
function showSP(){ document.getElementById('login').style.display='none'; document.getElementById('st-portal').style.display='flex'; }
function hideSP(){ document.getElementById('st-portal').style.display='none'; document.getElementById('login').style.display='flex'; }
function loginSt(){ let c=document.getElementById('sc').value.trim(); let s=DB.students.find(x=>x.code===c); if(s){ document.getElementById('st-portal').style.display='none'; document.getElementById('st-view').style.display='block'; document.getElementById('st-name').innerText='مرحباً، '+s.name; genStudentReport(s.id, 'st-grades'); } else tst('كود خطأ','e'); }
function closeModal(){ document.getElementById('modal').style.display='none'; }

document.querySelectorAll('.nb').forEach(b=>{
  b.addEventListener('click',()=>{
    document.querySelectorAll('.nb').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
    b.classList.add('active'); document.getElementById(b.dataset.tab).classList.add('active');
    if(b.dataset.tab==='att'){ document.getElementById('ad').valueAsDate=new Date(); setTimeout(() => { uAS(); rA(); }, 100); }
    if(b.dataset.tab==='dash') updateD();
    ref();
  });
});

function updateD(){
  document.getElementById('ds').innerText = DB.students.length;
  let today = new Date().toISOString().split('T')[0]; let ac = 0;
  for(let k in DB.attendance){ if(k.includes(today)) ac += Object.keys(DB.attendance[k]).length; }
  document.getElementById('da').innerText = ac;
}

function showStats(){
  let m = DB.students.filter(s=>s.gender==='ذكر'), f = DB.students.filter(s=>s.gender==='أنثى');
  let h = `<div class="row" style="margin-bottom:10px;"><div class="dash" style="background:#e0f2fe;"><h4>ذكور المدرسة</h4><h3>${m.length}</h3></div><div class="dash" style="background:#fce7f3;"><h4>إناث المدرسة</h4><h3>${f.length}</h3></div></div>`;
  h += `<h4 style="margin-bottom:5px; font-size:13px; color:var(--p);">حسب الصف:</h4>`;
  DB.levels.forEach(lv => {
    let lm = DB.students.filter(s=>s.level===lv && s.gender==='ذكر').length;
    let lf = DB.students.filter(s=>s.level===lv && s.gender==='أنثى').length;
    h += `<div class="li"><div><b>${lv}</b></div><div>ذكور: ${lm} | إناث: ${lf}</div></div>`;
  });
  h += `<h4 style="margin-top:10px; margin-bottom:5px; font-size:13px; color:var(--p);">حسب الشعبة:</h4>`;
  DB.sections.forEach(s => {
    let sm = DB.students.filter(x=>x.level===s.level && x.section===s.name && x.gender==='ذكر').length;
    let sf = DB.students.filter(x=>x.level===s.level && x.section===s.name && x.gender==='أنثى').length;
    h += `<div class="li"><div><b>${s.level} - ${s.name}</b></div><div>ذكور: ${sm} | إناث: ${sf}</div></div>`;
  });
  document.getElementById('stats-box').innerHTML = h;
}

function addL(){ let n=document.getElementById('nl').value.trim(); if(n && !DB.levels.includes(n)){ DB.levels.push(n); sv(); ref(); document.getElementById('nl').value=''; tst('تمت الإضافة'); } else tst('موجود','e'); }
function delL(n){ if(confirm('حذف؟')){ DB.levels=DB.levels.filter(x=>x!==n); DB.sections=DB.sections.filter(s=>s.level!==n); DB.students=DB.students.filter(s=>s.level!==n); sv(); init(); } }
function addS(){ let lv=document.getElementById('sl').value, n=document.getElementById('ns').value.trim(); if(!lv) return tst('اختر الصف','e'); if(n && !DB.sections.some(s=>s.level===lv && s.name===n)){ DB.sections.push({level:lv,name:n}); sv(); ref(); document.getElementById('ns').value=''; tst('تمت'); } else tst('موجودة','e'); }
function delS(l,n){ if(confirm('حذف؟')){ DB.sections=DB.sections.filter(s=>!(s.level===l && s.name===n)); DB.students=DB.students.filter(s=>!(s.level===l && s.section===n)); sv(); init(); } }
function uSS(){ let lv=document.getElementById('sl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('slist').innerHTML=f.map(s=>`<div class="li"><div><b>${s.name}</b></div><button class="ib d" onclick="delS('${s.level}','${s.name}')">🗑️</button></div>`).join(''); }
function uSF(){ let lv=document.getElementById('stl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('sts').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); rS(); }
function uGS(){ let lv=document.getElementById('gl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('gs').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); document.getElementById('gt').innerHTML=''; }
function uAS(){ let lv=document.getElementById('al').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('as').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); document.getElementById('att-list').innerHTML=''; }
function uACS(){ let lv=document.getElementById('acl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('acs').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); document.getElementById('act-list').innerHTML=''; }
function uLS(){ let lv=document.getElementById('ll').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('ls').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); document.getElementById('lvl-list').innerHTML=''; }
function uIS(){ let lv=document.getElementById('il').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('is').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); }
function uMS(){ let lv=document.getElementById('mvl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('mvs').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); }
function uRS(){ let lv=document.getElementById('rl').value; let f=DB.sections.filter(s=>s.level===lv); document.getElementById('rs').innerHTML='<option value="">-- الكل --</option>'+f.map(s=>`<option value="${s.name}">${s.name}</option>`).join(''); }

function importS(){
  let f=document.getElementById('ef').files[0], lv=document.getElementById('il').value, sec=document.getElementById('is').value, sg=document.getElementById('ig').value;
  if(!lv||!sec) return tst('اختر الصف والشعبة','e'); if(!f) return tst('اختر الملف','e');
  let r=new FileReader();
  r.onload=e=>{ try{ let w=XLSX.read(new Uint8Array(e.target.result),{type:'array'}); let rows=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1}); let ci=0; for(let i=0;i<rows[0].length;i++){ if(String(rows[0][i]).includes('اسم')||String(rows[0][i]).toLowerCase().includes('name')){ ci=i; break; } } let imp=0, skp=0; for(let i=1;i<rows.length;i++){ let rw=rows[i]; if(rw && rw[ci]){ let n=String(rw[ci]).trim(); if(n){ if(DB.students.some(x=>x.name===n)){ skp++; continue; } let g=(sg==='auto')?(rw[ci+1]||'ذكر'):sg; let par=rw[ci+2]||''; DB.students.push({id:nid(DB.students),name:n,level:lv,section:sec,code:sec+'-'+(1000+DB.students.length+1),gender:g,parentId:par}); imp++; } } } sv(); init(); document.getElementById('ef').value=''; tst(`استيراد ${imp} (مكرر: ${skp})`); }catch(err){ tst('خطأ','e'); } };
  r.readAsArrayBuffer(f);
}

function importG(){ document.getElementById('gfile').click(); }
function procG(input){
  let f=input.files[0]; if(!f) return;
  let s=document.getElementById('gs').value, sid=+document.getElementById('gsub').value;
  if(!s||!sid) return tst('اختر الشعبة والمادة أولاً','e');
  
  let r=new FileReader();
  r.onload=e=>{ 
    try{ 
      let w=XLSX.read(new Uint8Array(e.target.result),{type:'array'}); 
      let rows=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1}); 
      let imp=0; 
      
      for(let i=1;i<rows.length;i++){ 
        let rw=rows[i]; 
        if(!rw || !rw[0]) continue; 
        
        let name = String(rw[0]).trim();
        let st = DB.students.find(x=>x.name === name); 
        
        if(st){ 
          let ex=DB.grades.find(y=>y.student_id===st.id && y.subject_id===sid); 
          if(!ex){ ex={id:nid(DB.grades), student_id:st.id, subject_id:sid}; DB.grades.push(ex); } 
          
          ex.prep = +rw[1]||0;
          ex.book = +rw[2]||0;
          ex.cond = +rw[3]||0;
          ex.att = +rw[4]||0;
          ex.m1_w = +rw[5]||0;
          
          imp++; 
        } 
      } 
      sv(); 
      rG(); 
      input.value=''; 
      tst(`تم استيراد درجات ${imp} طالب بنجاح`); 
    }catch(err){ 
      tst('حدث خطأ في الملف','e'); 
    } 
  };
  r.readAsArrayBuffer(f);
}

function addSub(){ let n=document.getElementById('sn').value.trim(); if(n){ DB.subjects.push({id:nid(DB.subjects),name:n}); sv(); ref(); document.getElementById('sn').value=''; tst('تمت'); } }
function delSub(id){ if(confirm('حذف؟')){ DB.subjects=DB.subjects.filter(x=>x.id!==id); sv(); init(); } }
function delSt(id){ if(confirm('حذف؟')){ DB.students=DB.students.filter(x=>x.id!==id); sv(); rS(); } }
function clearAll(){ if(confirm('مسح الكل؟')){ if(prompt('اكتب "حذف":')==='حذف'){ DB={"students":[],"levels":[],"sections":[],"subjects":[{"id":1,"name":"اسلامية"},{"id":2,"name":"عربي"},{"id":3,"name":"انكليزي"},{"id":4,"name":"رياضيات"},{"id":5,"name":"احياء"},{"id":6,"name":"فيزياء"},{"id":7,"name":"كيمياء"},{"id":8,"name":"اجتماعيات"},{"id":9,"name":"اخلاقيه"}],"grades":[],"attendance":{},"activity":{},"adminPass":"admin"}; sv(); init(); tst('تم المسح'); } } }

function rS(){
  let q=document.getElementById('srch').value, lv=document.getElementById('stl').value, sec=document.getElementById('sts').value;
  let f=DB.students.filter(x=>(x.name.includes(q)||x.code.includes(q))&&(lv===''||x.level===lv)&&(sec===''||x.section===sec));
  let h=document.getElementById('stl-list');
  if(!f.length) return h.innerHTML='<p style="text-align:center;padding:15px;">لا يوجد</p>';
  h.innerHTML=f.map(x=>`<div class="li"><div style="display:flex;align-items:center;gap:8px;"><input type="checkbox" class="mvc" value="${x.id}" style="width:18px;height:18px;margin:0;"><div style="cursor:pointer;" onclick="openModal(${x.id})"><b style="text-decoration:underline;">${x.name}</b><small>${x.level} - ${x.section} | ${x.code} | ${x.gender||'-'}</small></div></div><div style="display:flex;gap:4px;"><button class="ib w" onclick="editS(${x.id})">⚙️</button><button class="ib d" onclick="delSt(${x.id})">🗑️</button></div></div>`).join('');
}

function openModal(id){
  let st=DB.students.find(x=>x.id===id); if(!st) return;
  genStudentReport(id, 'modal-content');
  document.getElementById('modal').style.display = 'flex';
}

function editS(id){
  let st=DB.students.find(x=>x.id===id); if(!st) return;
  let a=prompt(`1- تعديل الاسم\n2- تغيير الجنس\n3- ترقية لصف\nاكتب الرقم:`);
  if(a==='1'){ let n=prompt("الاسم:",st.name); if(n===null) return; st.name=n.trim(); sv(); rS(); tst('تم'); }
  else if(a==='2'){ let g=prompt("الجنس (ذكر/أنثى):",st.gender||'ذكر'); if(g===null) return; if(g!=='ذكر'&&g!=='أنثى') return tst('خطأ','e'); st.gender=g; sv(); rS(); tst('تم'); }
}

function printL(){
  let lv=document.getElementById('stl').value, sec=document.getElementById('sts').value;
  if(!lv||!sec) return tst('اختر','e');
  let st=DB.students.filter(x=>x.level===lv&&x.section===sec);
  if(!st.length) return tst('لا يوجد','e');
  let w=window.open('','','width=800,height=600');
  w.document.write(`<div style="text-align:center;font-family:Tahoma;padding:20px;"><h2>كشف (${lv} - ${sec})</h2><table border="1" style="width:100%;border-collapse:collapse;"><tr><th>م</th><th>الاسم</th><th>الكود</th><th>الجنس</th></tr>`);
  st.forEach((x,i)=>w.document.write(`<tr><td>${i+1}</td><td>${x.name}</td><td>${x.code}</td><td>${x.gender||'-'}</td></tr>`));
  w.document.write('</table></div>'); w.document.close(); w.print();
}

function moveSel(){
  let lv=document.getElementById('mvl').value, sec=document.getElementById('mvs').value;
  if(!lv||!sec) return tst('اختر','e');
  let chk=document.querySelectorAll('.mvc:checked'); if(!chk.length) return tst('اختر طلاباً','e');
  let c=0; chk.forEach(ch=>{ let st=DB.students.find(s=>s.id==ch.value); if(st){ st.level=lv; st.section=sec; st.code=sec+'-'+Math.floor(1000+Math.random()*9000); c++; } });
  sv(); rS(); tst(`نقل ${c}`);
}

function upgradeAll(){
  let fl=document.getElementById('ufl').value, tl=document.getElementById('utl').value;
  if(!fl||!tl||fl===tl) return tst('خطأ','e');
  if(!confirm(`ترقية الناجحين من ${fl} إلى ${tl}؟`)) return;
  let sts=DB.students.filter(s=>s.level===fl); if(!sts.length) return tst('لا يوجد','e');
  let c=0, yr=new Date().getFullYear();
  sts.forEach(st=>{
    let gs=DB.grades.filter(g=>g.student_id===st.id); let t=0,sc=0;
    gs.forEach(g=>{ let s=DB.subjects.find(x=>x.id===g.subject_id); if(s){ let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let final=(daily+(g.m1_w||0))/2; t+=final; sc++; } });
    if(sc>0&&(t/sc)>=50){ if(!st.archive) st.archive={}; st.archive[yr]=gs.map(g=>({...g})); st.level=tl; let as=DB.sections.filter(s=>s.level===tl).map(s=>s.name); if(!as.includes(st.section)) st.section=as.length?as[0]:'أ'; st.code=st.section+'-'+Math.floor(1000+Math.random()*9000); DB.grades=DB.grades.filter(g=>g.student_id!==st.id); c++; }
  });
  sv(); rS(); updateD(); tst(`ترقية ${c}`);
}

function init(){ ref(); rS(); updateD(); }
function ref(){
  let lo=DB.levels.map(x=>`<option value="${x}">${x}</option>`).join(''); let suo=DB.subjects.map(x=>`<option value="${x.id}">${x.name}</option>`).join('');
  document.getElementById('stl').innerHTML='<option value="">-- الكل --</option>'+lo;
  document.getElementById('sl').innerHTML='<option value="">-- اختر --</option>'+lo;
  document.getElementById('gl').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('al').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('acl').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('ll').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('rl').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('il').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('ufl').innerHTML='<option value="">-- من صف --</option>'+lo;
  document.getElementById('utl').innerHTML='<option value="">-- إلى صف --</option>'+lo;
  document.getElementById('mvl').innerHTML='<option value="">-- الصف --</option>'+lo;
  document.getElementById('mvl').onchange=uMS;
  document.getElementById('gsub').innerHTML='<option value="">-- المادة --</option>'+suo;
  document.getElementById('llist').innerHTML=DB.levels.map(x=>`<div class="li"><div><b>${x}</b></div><button class="ib d" onclick="delL('${x}')">🗑️</button></div>`).join('');
  document.getElementById('sublist').innerHTML=DB.subjects.map(x=>`<div class="li"><div><b>${x.name}</b></div><button class="ib d" onclick="delSub(${x.id})">🗑️</button></div>`).join('');
}

function autoS(sid,f,v,stId){ let e=DB.grades.find(y=>y.student_id===stId&&y.subject_id===sid); if(!e){ e={id:nid(DB.grades),student_id:stId,subject_id:sid}; DB.grades.push(e); } e[f]=v!==''?+v:null; sv(); rG(); }

function rG(){
  let s=document.getElementById('gs').value, sid=+document.getElementById('gsub').value, c=document.getElementById('gt');
  if(!s||!sid){ c.innerHTML=''; return; }
  let lv=document.getElementById('gl').value; let st=DB.students.filter(x=>x.section===s&&x.level===lv);
  if(!st.length){ c.innerHTML='<p style="text-align:center;">لا يوجد</p>'; return; }
  
  let h='<div style="overflow-x:auto;"><table><thead><tr><th>الطالب</th><th>تحضير (25)</th><th>دفتر (25)</th><th>سلوك (25)</th><th>حضور (25)</th><th>اليومي (100)</th><th>تحريري (100)</th><th>النهائية (100)</th><th>النتيجة</th></tr></thead><tbody>';
  st.forEach(x=>{
    let g=DB.grades.find(y=>y.student_id===x.id&&y.subject_id===sid)||{};
    let v=k=>g[k]!==undefined&&g[k]!==null?g[k]:'';
    let p = +g.prep || 0;
    let b = +g.book || 0;
    let cnd = +g.cond || 0;
    let at = +g.att || 0;
    let daily = p + b + cnd + at;
    let exam = +g.m1_w || 0;
    let final = (daily + exam) / 2;
    let stat = final >= 50 ? 'ناجح' : 'راسب';
    
    h+=`<tr>
    <td style="text-align:right;font-weight:bold; position:sticky; right:0; background:var(--card);">${x.name}</td>
    <td><input type="number" max="25" value="${v('prep')}" onchange="autoS(${sid},'prep',this.value,${x.id})" style="width:50px;padding:4px;margin:0;text-align:center;"></td>
    <td><input type="number" max="25" value="${v('book')}" onchange="autoS(${sid},'book',this.value,${x.id})" style="width:50px;padding:4px;margin:0;text-align:center;"></td>
    <td><input type="number" max="25" value="${v('cond')}" onchange="autoS(${sid},'cond',this.value,${x.id})" style="width:50px;padding:4px;margin:0;text-align:center;"></td>
    <td><input type="number" max="25" value="${v('att')}" onchange="autoS(${sid},'att',this.value,${x.id})" style="width:50px;padding:4px;margin:0;text-align:center;"></td>
    <td style="font-weight:bold;">${daily}</td>
    <td><input type="number" max="100" value="${v('m1_w')}" onchange="autoS(${sid},'m1_w',this.value,${x.id})" style="width:60px;padding:4px;margin:0;text-align:center;"></td>
    <td style="font-weight:bold;color:green;">${final.toFixed(1)}</td>
    <td style="font-weight:bold;color:${stat==='ناجح'?'green':'red'};">${stat}</td>
    </tr>`;
  });
  c.innerHTML=h+'</tbody></table></div>';
}

function rA(){
  let s=document.getElementById('as').value, d=document.getElementById('ad').value, q=document.getElementById('sa').value, lv=document.getElementById('al').value;
  if(!s||!d||!lv) return; let k=s+'_'+d; if(!DB.attendance[k]) DB.attendance[k]={};
  let st=DB.students.filter(x=>x.section===s&&x.level===lv&&x.name.includes(q));
  document.getElementById('att-list').innerHTML=st.map(x=>{
    let stt=DB.attendance[k][x.id]; let isA=stt===true||stt==='مريض'||stt==='عمل'||stt==='أخرى';
    let rt=stt===true?'بدون عذر (خصم 5)':(stt==='مريض'?'بعذر (مريض)':(stt==='عمل'?'بعذر (عمل)':(stt==='أخرى'?'بعذر (أخرى)':'حاضر')));
    return `<div class="li" style="flex-direction:column;align-items:flex-start;"><div style="display:flex;justify-content:space-between;width:100%;margin-bottom:5px;"><b>${x.name}</b><span style="color:${isA?'var(--e)':'var(--s)'};font-size:12px;font-weight:bold;">${rt}</span></div><div style="display:flex;gap:8px;align-items:center;width:100%;"><label style="font-size:12px;"><input type="checkbox" style="width:16px;height:16px;" ${isA?'checked':''} onchange="tA('${k}',${x.id},this.checked)"> غياب</label><select onchange="cA('${k}',${x.id},this.value)" style="padding:4px;font-size:11px;width:auto;flex:1;margin:0;display:${isA?'block':'none'};"><option value="true" ${stt===true?'selected':''}>بدون عذر</option><option value="مريض" ${stt==='مريض'?'selected':''}>مريض</option><option value="عمل" ${stt==='عمل'?'selected':''}>عمل</option><option value="أخرى" ${stt==='أخرى'?'selected':''}>أخرى</option></select></div></div>`;
  }).join('');
}
function tA(k,sid,chk){ if(!chk) delete DB.attendance[k][sid]; else DB.attendance[k][sid]='true'; sv(); rA(); }
function cA(k,sid,v){ DB.attendance[k][sid]=v; sv(); rA(); }
function calcA(sid){ let t=0,j=0; for(let k in DB.attendance){ let s=DB.attendance[k][sid]; if(s){ if(s===true||s==='true') t++; else j++; } } return {t,j}; }

function rAC(){
  let s=document.getElementById('acs').value, q=document.getElementById('sac').value, lv=document.getElementById('acl').value;
  if(!s||!lv) return; let st=DB.students.filter(x=>x.section===s&&x.level===lv&&x.name.includes(q));
  document.getElementById('act-list').innerHTML=st.map(x=>`<div class="li"><div><b>${x.name}</b></div><input type="number" id="ac-${x.id}" value="${DB.activity[x.id]||''}" style="width:70px;margin:0;padding:6px;text-align:center;"></div>`).join('');
  document.getElementById('sact').style.display='block';
}
function saveAC(){ let s=document.getElementById('acs').value; DB.students.filter(x=>x.section===s).forEach(x=>{ let v=document.getElementById(`ac-${x.id}`).value; if(v) DB.activity[x.id]=+v; else delete DB.activity[x.id]; }); sv(); tst('تم'); }

function rL(){
  let lv=document.getElementById('ll').value, sec=document.getElementById('ls').value, fl=document.getElementById('lf').value, el=document.getElementById('lvl-list');
  if(!lv||!sec){ el.innerHTML='<p style="text-align:center;">اختر</p>'; return; }
  let st=DB.students.filter(x=>x.level===lv&&x.section===sec); if(!st.length){ el.innerHTML='<p style="text-align:center;">لا يوجد</p>'; return; }
  let res=[];
  st.forEach(s=>{
    let gs=DB.grades.filter(g=>g.student_id===s.id); let t=0,c=0;
    gs.forEach(g=>{ let sb=DB.subjects.find(x=>x.id===g.subject_id); if(sb){ if(fl==='m1'){ let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let final=(daily+(g.m1_w||0))/2; t+=final; c++; } else if(fl==='m2'){ let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let final=(daily+(g.m2_w||0))/2; t+=final; c++; } else if(fl==='mid'){ t+=(g.mid||0); c++; } else if(fl==='final'){ t+=(g.final||0); c++; } else { let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let final=(daily+(g.m1_w||0))/2; t+=final; c++; } } });
    if(c>0){ let avg=t/c; let gr=''; if(avg>=90) gr='ممتاز'; else if(avg>=80) gr='جيد جداً'; else if(avg>=70) gr='جيد'; else if(avg>=60) gr='متوسط'; else if(avg>=50) gr='مقبول'; else gr='راسب'; res.push({n:s.name,a:avg.toFixed(1),g:gr,gn:s.gender||'-'}); }
  });
  res.sort((a,b)=>b.a-a.a);
  let males = res.filter(r=>r.gn==='ذكر'); let females = res.filter(r=>r.gn==='أنثى');
  let avgM = males.length ? (males.reduce((a,b)=>a+parseFloat(b.a),0)/males.length).toFixed(1) : 0;
  let avgF = females.length ? (females.reduce((a,b)=>a+parseFloat(b.a),0)/females.length).toFixed(1) : 0;
  let h = `<div class="row" style="margin-bottom:10px;"><div class="dash" style="background:#e0f2fe;"><h4>معدل الذكور (${males.length})</h4><h3>${avgM}%</h3></div><div class="dash" style="background:#fce7f3;"><h4>معدل الإناث (${females.length})</h4><h3 style="color:#db2777;">${avgF}%</h3></div></div>`;
  h += '<table><thead><tr><th>#</th><th>الاسم</th><th>الجنس</th><th>المعدل</th><th>التقدير</th></tr></thead><tbody>';
  res.forEach((r,i)=>{ let col=r.g==='ممتاز'?'var(--s)':(r.g==='راسب'?'var(--e)':'var(--p)'); h+=`<tr><td>${i+1}</td><td style="font-weight:bold;">${r.n}</td><td>${r.gn}</td><td>${r.a}%</td><td style="color:${col};font-weight:bold;">${r.g}</td></tr>`; });
  h+='</tbody></table>'; el.innerHTML=h;
}

function genRep(){
  let lv=document.getElementById('rl').value, sec=document.getElementById('rs').value, type=document.getElementById('rtype').value, el=document.getElementById('rbox');
  if(!lv||!sec) return tst('اختر الصف والشعبة','e');
  let st=DB.students.filter(x=>x.level===lv&&x.section===sec);
  if(!st.length) return tst('لا يوجد','e');
  let h='';
  if(type==='grades'){
    h='<table><thead><tr><th>الطالب</th><th>ف1</th><th>ف2</th><th>السعي</th><th>النهائي</th><th>النتيجة</th></tr></thead><tbody>';
    st.forEach(s=>{
      let gs=DB.grades.filter(g=>g.student_id===s.id); let t=0,c=0;
      gs.forEach(g=>{ let sb=DB.subjects.find(x=>x.id===g.subject_id); if(sb){ let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let final=(daily+(g.m1_w||0))/2; t+=final; c++; } });
      let avg=c?(t/c):0; h+=`<tr><td>${s.name}</td><td>—</td><td>—</td><td>${avg.toFixed(1)}</td><td>—</td><td>${avg>=50?'ناجح':'راسب'}</td></tr>`;
    });
    h+='</tbody></table>';
  } else if(type==='att'){
    h='<table><thead><tr><th>الطالب</th><th>بدون عذر</th><th>بعذر</th></tr></thead><tbody>';
    st.forEach(s=>{ let a=calcA(s.id); h+=`<tr><td>${s.name}</td><td>${a.t}</td><td>${a.j}</td></tr>`; });
    h+='</tbody></table>';
  } else if(type==='act'){
    h='<table><thead><tr><th>الطالب</th><th>النشاط</th></tr></thead><tbody>';
    st.forEach(s=>{ h+=`<tr><td>${s.name}</td><td>${DB.activity[s.id]||0}</td></tr>`; });
    h+='</tbody></table>';
  }
  el.innerHTML=h;
}

function genStudentReport(sid, box){
  let st=DB.students.find(x=>x.id===sid); if(!st) return;
  let gs=DB.grades.filter(x=>x.student_id===sid); let a=calcA(sid); let act=DB.activity[sid]||0;
  let rel=DB.students.filter(s=>s.parentId===st.parentId&&s.id!==st.id&&st.parentId); let rt=rel.length?rel.map(r=>r.name).join('، '):'لا يوجد';
  let h=`<div class="rc"><div class="rh"><h3>${st.name}</h3><div class="si"><div>${st.level} - ${st.section}</div><div>${st.code}</div><div>${st.gender||'-'}</div><div style="color:var(--e);">غياب: ${a.t} | ${a.j}</div><div style="color:var(--s);">نشاط: ${act}</div><div style="width:100%;">👨‍👩‍👦 ${rt}</div></div></div><div class="gg">`;
  if(!gs.length) h+='<p style="text-align:center;">لا توجد</p>';
  gs.forEach(g=>{
    let sb=DB.subjects.find(x=>x.id===g.subject_id); if(sb){
      let daily = (+g.prep||0) + (+g.book||0) + (+g.cond||0) + (+g.att||0);
      let exam = g.m1_w || 0;
      let final = (daily + exam) / 2;
      let stat = final >= 50 ? 'ناجح' : 'راسب';
      h+=`<div class="sc"><div class="st">${sb.name}</div><div class="gr">تحضير: <span>${g.prep||0}</span></div><div class="gr">دفتر: <span>${g.book||0}</span></div><div class="gr">سلوك: <span>${g.cond||0}</span></div><div class="gr">حضور: <span>${g.att||0}</span></div><div class="gr" style="color:var(--p);font-weight:bold;">اليومي: <span>${daily}</span></div><div class="gr">تحريري: <span>${exam}</span></div><div class="gr" style="color:var(--w);font-weight:bold;">النهائية: <span>${final.toFixed(1)}</span></div><div class="fb ${stat==='راسب'?'f':''}">${final.toFixed(1)} (${stat})</div></div>`;
    }
  });
  h+='</div></div>'; document.getElementById(box).innerHTML=h;
}

function expCSV(){
  let csv='\uFEFFالاسم,الجنس,الصف,الشعبة,المادة,تحضير,دفتر,سلوك,حضور,اليومي,تحريري,النهائية,النتيجة\n';
  DB.students.forEach(st=>{ let gs=DB.grades.filter(g=>g.student_id===st.id); gs.forEach(g=>{ let sb=DB.subjects.find(x=>x.id===g.subject_id); if(sb){ let daily=((g.prep||0)+(g.book||0)+(g.cond||0)+(g.att||0)); let exam=g.m1_w||0; let final=(daily+exam)/2; let stt=final>=50?'ناجح':'راسب'; csv+=`"${st.name}","${st.gender||'-'}","${st.level}","${st.section}","${sb.name}",${g.prep||0},${g.book||0},${g.cond||0},${g.att||0},${daily},${exam},${final.toFixed(1)},"${stt}"\n`; } }); });
  let blob=new Blob([csv],{type:'text/csv;charset=utf-8;'}); let a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='درجات.csv'; a.click(); tst('تم');
}
</script>
</body>
</html>
