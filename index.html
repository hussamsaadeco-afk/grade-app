<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>نظام إدارة الدرجات - النسخة الاحترافية</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #eef2f6; color: #1f2937; padding-bottom: 70px; min-height: 100vh; }
header { background: #1e3a8a; color: white; padding: 15px; text-align: center; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 10px rgba(0,0,0,.2); }
header h1 { font-size: 18px; }
main { padding: 10px; }
.tab { display: none; }
.tab.active { display: block; animation: fadeIn .2s; }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
.card { background: white; border-radius: 12px; padding: 15px; margin-bottom: 10px; box-shadow: 0 2px 6px rgba(0,0,0,.05); }
.card h2 { font-size: 15px; color: #1e3a8a; margin-bottom: 12px; border-bottom: 1px solid #eee; padding-bottom: 8px; }
input, select, button { width: 100%; padding: 12px; margin-bottom: 8px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 15px; font-family: inherit; }
button { cursor: pointer; font-weight: bold; border: none; }
.btn-primary { background: #16a34a; color: white; }
.btn-info { background: #3b82f6; color: white; }
.btn-warning { background: #f59e0b; color: white; }
.btn-danger { background: #dc2626; color: white; }
.row { display: flex; gap: 8px; }
.row > * { flex: 1; }
.list { margin-top: 10px; }
.item { background: #f8fafc; border-right: 4px solid #3b82f6; padding: 10px; margin-bottom: 8px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; }
.item-info strong { display: block; font-size: 14px; }
.item-info small { color: #64748b; font-size: 12px; }
.icon-btn { width: auto; padding: 6px 10px; margin: 0; background: #ef4444; color: white; border-radius: 6px; }
.bottom-nav { position: fixed; bottom: 0; left: 0; right: 0; background: white; display: flex; justify-content: space-around; padding: 6px 0; box-shadow: 0 -2px 10px rgba(0,0,0,.1); z-index: 100; }
.nav-btn { background: none; display: flex; flex-direction: column; align-items: center; color: #94a3b8; font-size: 20px; width: auto; padding: 5px 10px; margin: 0; }
.nav-btn span { font-size: 10px; margin-top: 2px; font-weight: bold; }
.nav-btn.active { color: #1e3a8a; background: #eff6ff; border-radius: 8px; }
table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
th, td { border: 1px solid #cbd5e1; padding: 8px; text-align: center; }
th { background: #f1f5f9; color: #1e3a8a; }
.toast { position: fixed; bottom: 80px; left: 50%; transform: translateX(-50%) translateY(100px); background: #111827; color: white; padding: 10px 20px; border-radius: 20px; font-size: 13px; opacity: 0; transition: all .3s; z-index: 200; }
.toast.show { transform: translateX(-50%) translateY(0); opacity: 1; }
.toast.error { background: #dc2626; }
.toast.success { background: #16a34a; }
</style>
</head>
<body>

<header><h1>🏫 نظام إدارة الدرجات الاحترافي</h1></header>

<nav class="bottom-nav">
  <button class="nav-btn active" data-tab="students">👨‍🎓<span>الطلاب</span></button>
  <button class="nav-btn" data-tab="subjects">📚<span>المواد</span></button>
  <button class="nav-btn" data-tab="grades">📝<span>الدرجات</span></button>
  <button class="nav-btn" data-tab="reports">📄<span>الكشوف</span></button>
  <button class="nav-btn" data-tab="backup">💾<span>النسخ</span></button>
</nav>

<main>
  <!-- تبويب الطلاب -->
  <section id="students" class="tab active">
    <div class="card">
      <h2>📥 استيراد الطلاب من الإكسل</h2>
      <p style="font-size:12px; color:#666; margin-bottom:10px;">ملف الإكسل (CSV) يجب أن يحتوي على الأعمدة التالية بالترتيب: (الاسم، الرقم الوطني، الصف، الشعبة)</p>
      <input type="file" id="excel-file" accept=".csv" style="border:none; padding:5px;">
      <button class="btn-info" onclick="importStudents()">استيراد الأسماء</button>
    </div>
    
    <div class="card">
      <h2>➕ إضافة طالب يدوياً</h2>
      <input id="s-name" placeholder="اسم الطالب">
      <div class="row">
        <input id="s-nid" placeholder="الرقم الوطني">
        <input id="s-level" placeholder="الصف">
      </div>
      <input id="s-section" placeholder="الشعبة">
      <button class="btn-primary" onclick="addStudent()">إضافة طالب</button>
    </div>

    <div class="card">
      <input id="search-student" placeholder="🔍 بحث بالاسم أو الشعبة..." oninput="renderStudents()">
      <div id="students-list" class="list"></div>
    </div>
  </section>

  <!-- تبويب المواد -->
  <section id="subjects" class="tab">
    <div class="card">
      <h2>➕ إضافة مادة جديدة</h2>
      <input id="sub-name" placeholder="اسم المادة">
      <div class="row">
        <input id="sub-max" placeholder="الدرجة العظمى" value="100" type="number">
        <input id="sub-pass" placeholder="درجة النجاح" value="50" type="number">
      </div>
      <button class="btn-primary" onclick="addSubject()">إضافة المادة</button>
    </div>
    <div class="card"><div id="subjects-list" class="list"></div></div>
  </section>

  <!-- تبويب الدرجات -->
  <section id="grades" class="tab">
    <div class="card">
      <h2>📝 إدخال الدرجات</h2>
      <select id="g-subject"></select>
      <select id="g-semester">
        <option>الفصل الأول</option>
        <option>الفصل الثاني</option>
      </select>
      <select id="g-section" onchange="renderGradeTable()">
        <option value="">-- اختر الشعبة --</option>
      </select>
      <button class="btn-info" onclick="renderGradeTable()">عرض الطلاب</button>
      
      <div id="grade-table-container" style="overflow-x:auto; margin-top:10px;">
        <!-- الجدول سيظهر هنا -->
      </div>
      <button class="btn-primary" id="save-grades-btn" style="display:none; margin-top:10px;" onclick="saveAllGrades()">💾 حفظ كل الدرجات</button>
    </div>
  </section>

  <!-- تبويب التقارير -->
  <section id="reports" class="tab">
    <div class="card">
      <h2>📄 كشف درجات طالب</h2>
      <select id="report-student" onchange="generateReportCard()"></select>
      <div id="report-card" style="margin-top:15px;"></div>
      <button class="btn-info" onclick="window.print()" style="margin-top:10px;">🖨️ طباعة الكشف</button>
    </div>
  </section>

  <!-- تبويب النسخ الاحتياطي -->
  <section id="backup" class="tab">
    <div class="card">
      <h2>💾 النسخ الاحتياطي والاستعادة</h2>
      <p style="font-size:13px; color:#555; margin-bottom:10px;">احفظ نسخة من بياناتك أو استعدها. (يفضل حفظ نسخة أسبوعياً).</p>
      <button class="btn-info" onclick="exportBackup()">📤 تصدير نسخة (Backup JSON)</button>
      <input type="file" id="restore-file" accept=".json" style="margin-top:10px; border:none;">
      <button class="btn-warning" onclick="importBackup()">📥 استعادة من ملف</button>
      <button class="btn-danger" onclick="clearAllData()" style="margin-top:20px;">🗑️ مسح كل البيانات</button>
    </div>
  </section>
</main>

<div id="toast" class="toast"></div>

<script>
let DB = JSON.parse(localStorage.getItem('proGradeDB') || '{"students":[],"subjects":[],"grades":[]}');
function save() { localStorage.setItem('proGradeDB', JSON.stringify(DB)); }
function toast(msg, type='success') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => t.className = 'toast ' + type, 2500);
}
function nextId(arr) { return arr.length ? Math.max(...arr.map(x => x.id)) + 1 : 1; }

// التنقل
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    if (btn.dataset.tab === 'grades') refreshGradeSelects();
    if (btn.dataset.tab === 'reports') refreshReportSelects();
  });
});

// --- الطلاب ---
function importStudents() {
  const file = document.getElementById('excel-file').files[0];
  if (!file) return toast('اختر ملف إكسل (CSV)', 'error');
  
  const reader = new FileReader();
  reader.onload = function(e) {
    const text = e.target.result;
    const rows = text.split('\n');
    let imported = 0;
    
    // تخطي السطر الأول (العناوين) إذا كان يحتوي على كلمات مثل "الاسم"
    let start = (rows[0].includes('الاسم') || rows[0].includes('Name')) ? 1 : 0;
    
    for (let i = start; i < rows.length; i++) {
      const cols = rows[i].split(',');
      if (cols.length >= 2 && cols[0].trim() !== '') {
        DB.students.push({
          id: nextId(DB.students),
          name: cols[0].trim().replace(/"/g, ''),
          national_id: (cols[1] || '').trim().replace(/"/g, ''),
          grade_level: (cols[2] || '').trim().replace(/"/g, ''),
          section: (cols[3] || '').trim().replace(/"/g, '')
        });
        imported++;
      }
    }
    save(); renderStudents();
    toast(`تم استيراد ${imported} طالب بنجاح ✅`);
  };
  reader.readAsText(file, 'UTF-8');
}

function addStudent() {
  const name = document.getElementById('s-name').value.trim();
  if (!name) return toast('الاسم مطلوب', 'error');
  DB.students.push({ id: nextId(DB.students), name, 
    national_id: document.getElementById('s-nid').value, 
    grade_level: document.getElementById('s-level').value, 
    section: document.getElementById('s-section').value });
  save();
  ['s-name','s-nid','s-level','s-section'].forEach(id => document.getElementById(id).value = '');
  renderStudents(); toast('تمت الإضافة');
}

function renderStudents() {
  const q = document.getElementById('search-student').value.trim();
  const list = document.getElementById('students-list');
  const filtered = DB.students.filter(s => s.name.includes(q) || s.section.includes(q));
  if (!filtered.length) { list.innerHTML = '<div style="text-align:center; color:#999; padding:10px;">لا يوجد طلاب</div>'; return; }
  
  list.innerHTML = filtered.map(s => `
    <div class="item">
      <div class="item-info">
        <strong>${s.name}</strong>
        <small>${s.national_id || '—'} • ${s.grade_level || '—'} - ${s.section || '—'}</small>
      </div>
      <button class="icon-btn" onclick="deleteStudent(${s.id})">🗑️</button>
    </div>`).join('');
}

function deleteStudent(id) {
  if (!confirm('حذف الطالب؟')) return;
  DB.students = DB.students.filter(s => s.id !== id);
  DB.grades = DB.grades.filter(g => g.student_id !== id);
  save(); renderStudents(); toast('تم الحذف');
}

// --- المواد ---
function addSubject() {
  const name = document.getElementById('sub-name').value.trim();
  if (!name) return toast('اسم المادة مطلوب', 'error');
  DB.subjects.push({ id: nextId(DB.subjects), name, 
    max: +document.getElementById('sub-max').value || 100, 
    pass: +document.getElementById('sub-pass').value || 50 });
  save(); document.getElementById('sub-name').value = '';
  renderSubjects(); toast('تمت إضافة المادة');
}

function renderSubjects() {
  const list = document.getElementById('subjects-list');
  if (!DB.subjects.length) { list.innerHTML = '<div style="text-align:center; color:#999; padding:10px;">لا توجد مواد</div>'; return; }
  list.innerHTML = DB.subjects.map(s => `
    <div class="item">
      <div class="item-info">
        <strong>${s.name}</strong>
        <small>العظمى: ${s.max} • النجاح: ${s.pass}</small>
      </div>
      <button class="icon-btn" onclick="deleteSubject(${s.id})">🗑️</button>
    </div>`).join('');
}

function deleteSubject(id) {
  if (!confirm('حذف المادة؟')) return;
  DB.subjects = DB.subjects.filter(s => s.id !== id);
  DB.grades = DB.grades.filter(g => g.subject_id !== id);
  save(); renderSubjects(); toast('تم الحذف');
}

// --- إدخال الدرجات الجماعي ---
function refreshGradeSelects() {
  document.getElementById('g-subject').innerHTML = '<option value="">اختر المادة</option>' + 
    DB.subjects.map(s => `<option value="${s.id}">${s.name}</option>`).join('');
  
  const sections = [...new Set(DB.students.map(s => s.section).filter(s => s))];
  document.getElementById('g-section').innerHTML = '<option value="">-- اختر الشعبة --</option>' + 
    sections.map(s => `<option value="${s}">${s}</option>`).join('');
}

let currentGradeData = [];
function renderGradeTable() {
  const subId = +document.getElementById('g-subject').value;
  const sec = document.getElementById('g-section').value;
  const sem = document.getElementById('g-semester').value;
  const container = document.getElementById('grade-table-container');
  const saveBtn = document.getElementById('save-grades-btn');

  if (!subId || !sec) { container.innerHTML = ''; saveBtn.style.display = 'none'; return; }

  const students = DB.students.filter(s => s.section === sec);
  if (!students.length) { container.innerHTML = '<p style="text-align:center;">لا يوجد طلاب في هذه الشعبة</p>'; saveBtn.style.display = 'none'; return; }

  let html = '<table><thead><tr><th>الطالب</th><th>الدرجة</th></tr></thead><tbody>';
  students.forEach(s => {
    const existing = DB.grades.find(g => g.student_id === s.id && g.subject_id === subId && g.semester === sem);
    const val = existing ? existing.grade : '';
    html += `<tr>
      <td style="text-align:right;">${s.name}</td>
      <td><input type="number" id="grade-${s.id}" value="${val}" style="margin:0; padding:5px; text-align:center;" inputmode="decimal"></td>
    </tr>`;
  });
  html += '</tbody></table>';
  container.innerHTML = html;
  saveBtn.style.display = 'block';
  currentGradeData = { subId, sem };
}

function saveAllGrades() {
  const { subId, sem } = currentGradeData;
  const sec = document.getElementById('g-section').value;
  const students = DB.students.filter(s => s.section === sec);
  
  students.forEach(s => {
    const input = document.getElementById(`grade-${s.id}`);
    if (input && input.value !== '') {
      const grade = +input.value;
      const existing = DB.grades.find(g => g.student_id === s.id && g.subject_id === subId && g.semester === sem);
      if (existing) existing.grade = grade;
      else DB.grades.push({ id: nextId(DB.grades), student_id: s.id, subject_id: subId, semester: sem, grade });
    }
  });
  save(); toast('تم حفظ درجات الشعبة بنجاح ✅');
}

// --- التقارير ---
function refreshReportSelects() {
  document.getElementById('report-student').innerHTML = '<option value="">-- اختر طالباً --</option>' + 
    DB.students.map(s => `<option value="${s.id}">${s.name}</option>`).join('');
  document.getElementById('report-card').innerHTML = '';
}

function generateReportCard() {
  const sid = +document.getElementById('report-student').value;
  const container = document.getElementById('report-card');
  if (!sid) { container.innerHTML = ''; return; }
  
  const student = DB.students.find(s => s.id === sid);
  const grades = DB.grades.filter(g => g.student_id === sid);
  
  let html = `<div style="border:1px solid #ccc; padding:15px; border-radius:8px; background:#fff;">
    <h3 style="text-align:center; color:#1e3a8a; margin-bottom:10px;">كشف درجات الطالب</h3>
    <p><strong>الاسم:</strong> ${student.name}</p>
    <p><strong>الصف:</strong> ${student.grade_level} - ${student.section}</p>
    <table style="margin-top:10px;">
      <thead><tr><th>المادة</th><th>الفصل</th><th>الدرجة</th><th>النتيجة</th></tr></thead>
      <tbody>`;
      
  let totalPct = 0; let count = 0;
  grades.forEach(g => {
    const sub = DB.subjects.find(s => s.id === g.subject_id);
    if (sub) {
      const ok = g.grade >= sub.pass;
      totalPct += (g.grade / sub.max) * 100;
      count++;
      html += `<tr>
        <td>${sub.name}</td>
        <td>${g.semester}</td>
        <td>${g.grade} / ${sub.max}</td>
        <td style="color:${ok?'green':'red'}">${ok?'ناجح':'راسب'}</td>
      </tr>`;
    }
  });
  
  const avg = count ? (totalPct / count).toFixed(1) : 0;
  html += `</tbody></table>
    <h4 style="margin-top:15px; text-align:center; background:#f0f2f5; padding:10px; border-radius:6px;">المعدل العام: ${avg}%</h4>
  </div>`;
  container.innerHTML = html;
}

// --- النسخ الاحتياطي ---
function exportBackup() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(DB));
  const a = document.createElement('a');
  a.href = dataStr;
  a.download = "نسخة_احتياطية_" + new Date().toLocaleDateString('ar-EG') + ".json";
  a.click();
  toast('تم تصدير النسخة');
}

function importBackup() {
  const file = document.getElementById('restore-file').files[0];
  if (!file) return toast('اختر ملف النسخة الاحتياطية', 'error');
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      if (data.students && data.subjects && data.grades) {
        DB = data; save();
        renderStudents(); renderSubjects(); refreshGradeSelects();
        toast('تم استعادة البيانات بنجاح ✅');
      } else { toast('ملف غير صالح', 'error'); }
    } catch(err) { toast('خطأ في قراءة الملف', 'error'); }
  };
  reader.readAsText(file);
}

function clearAllData() {
  if (confirm('هل أنت متأكد من مسح كل البيانات؟ لا يمكن التراجع!')) {
    DB = {students:[], subjects:[], grades:[]}; save();
    renderStudents(); renderSubjects(); refreshGradeSelects();
    toast('تم مسح البيانات');
  }
}

// التشغيل
renderStudents();
renderSubjects();
</script>
</body>
</html>
