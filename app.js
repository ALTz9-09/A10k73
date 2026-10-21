'use strict';
const $ = id => document.getElementById(id);
let selected = null;
let importedNames = [];
let selectedDesks = [];
let randomAssignments = [];
let animationToken = 0;
const seats = [];
const desks = [];

for (let d = 0; d < 24; d++) {
  const desk = document.createElement('div');
  desk.className = 'desk';
  desks.push(desk);
  for (let s = 0; s < 2; s++) {
    const i = d * 2 + s;
    const el = document.createElement('textarea');
    el.className = 'seat';
    el.maxLength = 70;
    el.setAttribute('aria-label', `Hàng ${Math.floor(d / 4) + 1}, tổ ${d % 4 + 1}, chỗ ${s + 1}`);
    el.addEventListener('input', updateCount);
    el.addEventListener('click', () => {
      if (!$('swap').checked) return;
      if (selected === null) { selected = i; el.classList.add('selected') }
      else {
        [seats[selected].value, el.value] = [el.value, seats[selected].value];
        resetSelection(); updateCount(); syncCurrentResults();
      }
    });
    desk.append(el); seats.push(el);
  }
  $('desks').append(desk);
}

function resetSelection() { selected = null; seats.forEach(s => s.classList.remove('selected')) }
function updateCount() { $('occupied').textContent = `${seats.filter(s => s.value.trim()).length} / 48 chỗ` }
function updateHeader() {
  $('title').textContent = `SƠ ĐỒ LỚP ${$('className').value.trim() || '…'} NĂM HỌC: ${$('year').value.trim() || '…'}`;
  $('subtitle').textContent = `GVCN: ${$('teacher').value.trim() || '…'}`;
}
['teacher', 'className', 'year'].forEach(id => $(id).addEventListener('input', updateHeader));
$('swap').onchange = () => { resetSelection(); seats.forEach(s => s.readOnly = $('swap').checked) };
$('clear').onclick = () => {
  if (seats.some(s => s.value) && confirm('Xóa toàn bộ tên trên sơ đồ?')) {
    seats.forEach(s => s.value = ''); resetArrangementState(); randomAssignments = []; $('export-random').disabled = true;
    resetSelection(); updateCount();
  }
};
$('print').onclick = () => window.print();

// Mẫu cố định: trang đầu tiên, cột A, tiêu đề ở A1.
$('template').onclick = () => {
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; return; }
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Họ và tên'], ...Array.from({ length: 48 }, () => [''])]);
  sheet['!cols'] = [{ wch: 38 }];
  XLSX.utils.book_append_sheet(book, sheet, 'Danh sách');
  XLSX.writeFile(book, 'danh-sach-hoc-sinh.xlsx');
};
function readStudentNames(book) {
  const sheet = book.Sheets[book.SheetNames[0]];
  if (!sheet || String(sheet.A1?.v || '').normalize('NFC').trim().toLocaleLowerCase('vi') !== 'họ và tên')
    throw Error('File chưa đúng mẫu. Hãy tải Excel mẫu và điền tên vào cột Họ và tên.');
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, range: 0 });
  const names = rows.slice(1).map(row => String(row[0] ?? '').trim()).filter(Boolean);
  if (names.length > 48) throw Error(`File có ${names.length} học sinh. Sơ đồ chỉ có tối đa 48 chỗ.`);
  if (!names.length) throw Error('File chưa có tên học sinh.');
  if (names.some(name => name.length > 70)) throw Error('Có họ tên dài hơn 70 ký tự. Vui lòng kiểm tra lại file.');
  return names;
}
$('file').onchange = async () => {
  const input = $('file'), file = input.files[0]; if (!file) return;
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; input.value = ''; return; }
  input.disabled = true; $('status').textContent = 'Đang đọc file…';
  try {
    if (file.size > 10 * 1024 * 1024) throw Error('Vui lòng chọn file nhỏ hơn 10 MB.');
    let book;
    try { book = XLSX.read(await file.arrayBuffer(), { type: 'array' }); }
    catch { throw Error('Không đọc được file. Vui lòng dùng Excel mẫu.'); }
    importedNames = readStudentNames(book);
    randomAssignments = []; selectedDesks = [];
    seats.forEach((seat, i) => seat.value = importedNames[i] || '');
    resetArrangementState();
    $('random').disabled = false; $('export-random').disabled = true;
    resetSelection(); updateCount();
    $('status').textContent = `Đã nhận ${importedNames.length} học sinh.`;
  } catch (error) { $('status').textContent = error.message; }
  finally { input.disabled = false; input.value = ''; }
};

// Chọn bàn theo thứ tự bấm. Thứ tự này chính là số bàn hiển thị khi quay random.
const pickerButtons = [];
for (let d = 0; d < 24; d++) {
  const button = document.createElement('button');
  button.type = 'button'; button.className = 'picker-desk';
  button.dataset.desk = String(d);
  button.setAttribute('aria-label', `Hàng ${Math.floor(d / 4) + 1}, Tổ ${d % 4 + 1}`);
  button.onclick = () => toggleDesk(d);
  pickerButtons.push(button); $('desk-picker').append(button);
}
// Sơ đồ thu nhỏ để nhìn tên học sinh bay vào đúng bàn trong lúc random.
const liveSection = document.createElement('section');
liveSection.className = 'random-live-section';
const liveTitle = document.createElement('h3'); liveTitle.textContent = 'Sơ đồ đang xếp chỗ';
const liveHeadings = document.createElement('div'); liveHeadings.className = 'live-headings';
['Tổ 1', 'Tổ 2', 'Tổ 3', 'Tổ 4'].forEach(label => { const span = document.createElement('span'); span.textContent = label; liveHeadings.append(span) });
const liveLayout = document.createElement('div'); liveLayout.className = 'random-live-layout';
const liveDesks = [];
for (let d = 0; d < 24; d++) {
  const desk = document.createElement('div'); desk.className = 'live-desk';
  const badge = document.createElement('b'); badge.className = 'live-desk-number';
  const left = document.createElement('span'), right = document.createElement('span'); left.className = right.className = 'live-seat';
  desk.append(badge, left, right); liveDesks.push(desk); liveLayout.append(desk);
}
liveSection.append(liveTitle, liveHeadings, liveLayout);
$('random-run-view').insertBefore(liveSection, $('assignment-list'));

function physicalDeskName(d) { return `Tổ ${d % 4 + 1} · hàng ${Math.floor(d / 4) + 1}` }
function refreshPicker() {
  pickerButtons.forEach((button, d) => {
    const order = selectedDesks.indexOf(d);
    button.classList.toggle('chosen', order >= 0);
    button.innerHTML = order >= 0 ? `<small>Bàn</small><strong>${order + 1}</strong>` : `<small>${Math.floor(d / 4) + 1}</small><span>Chọn</span>`;
    button.title = physicalDeskName(d);
  });
  const places = selectedDesks.length * 2;
  $('selected-count').textContent = selectedDesks.length ? `Đã chọn ${selectedDesks.length} bàn · ${places} chỗ` : 'Chưa chọn bàn';
  $('required-count').textContent = `Cần ${importedNames.length} chỗ`;
  $('random-error').textContent = '';
}
function toggleDesk(d) {
  const i = selectedDesks.indexOf(d);
  if (i >= 0) selectedDesks.splice(i, 1); else selectedDesks.push(d);
  refreshPicker();
}
function openRandom() {
  if (!importedNames.length) { $('status').textContent = 'Hãy tải danh sách Excel trước khi random.'; return; }
  selectedDesks = []; refreshPicker();
  $('random-select-view').hidden = false; $('random-run-view').hidden = true;
  $('random-modal').hidden = false; document.body.classList.add('modal-open');
}
function closeRandom() {
  animationToken++; $('random-modal').hidden = true; document.body.classList.remove('modal-open');
}
$('random').onclick = openRandom;
$('random-close').onclick = closeRandom;
$('random-modal').onclick = e => { if (e.target === $('random-modal') && !$('random-select-view').hidden) closeRandom() };
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('random-modal').hidden && !$('random-select-view').hidden) closeRandom() });
$('select-all').onclick = () => { selectedDesks = Array.from({ length: 24 }, (_, i) => i); refreshPicker() };

function shuffle(values) {
  const out = values.slice();
  for (let i = out.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1));[out[i], out[j]] = [out[j], out[i]] }
  return out;
}
function makeAssignments() {
  const slots = [];
  selectedDesks.forEach((desk, deskOrder) => {
    slots.push({ seat: desk * 2, deskNumber: deskOrder + 1, physicalDesk: desk });
    slots.push({ seat: desk * 2 + 1, deskNumber: deskOrder + 1, physicalDesk: desk });
  });
  const randomSlots = shuffle(slots);
  return importedNames.map((name, i) => ({ name, ...randomSlots[i] }));
}
function prepareLiveLayout() {
  liveDesks.forEach((desk, d) => {
    const order = selectedDesks.indexOf(d);
    desk.classList.toggle('active', order >= 0); desk.classList.remove('complete-celebration');
    desk.querySelector('.live-desk-number').textContent = order >= 0 ? `Bàn ${order + 1}` : '';
    desk.querySelectorAll('.live-seat').forEach(seat => { seat.textContent = ''; seat.classList.remove('filled') });
  });
}
function placeNameInLiveDesk(item) {
  const slot = liveDesks[item.physicalDesk].querySelectorAll('.live-seat')[item.seat % 2];
  slot.textContent = item.name; slot.title = item.name; slot.classList.add('filled');
}
async function flyNameToDesk(item, token) {
  const source = $('draw-name'), target = liveDesks[item.physicalDesk], targetSlot = target.querySelectorAll('.live-seat')[item.seat % 2];
  const from = source.getBoundingClientRect(), to = targetSlot.getBoundingClientRect();
  const flyer = document.createElement('div'); flyer.className = 'flying-student'; flyer.textContent = item.name;
  flyer.style.left = `${from.left}px`; flyer.style.top = `${from.top}px`; flyer.style.width = `${Math.min(Math.max(from.width, 170), 300)}px`;
  document.body.append(flyer);
  const dx = to.left + to.width / 2 - (from.left + Math.min(Math.max(from.width, 170), 300) / 2);
  const dy = to.top + to.height / 2 - (from.top + from.height / 2);
  const keyframes = [
    { transform: 'translate(0,0) scale(1)', opacity: 1, offset: 0 },
    { transform: `translate(${dx * .48}px,${dy * .48 - 85}px) scale(1.08)`, opacity: 1, offset: .52 },
    { transform: `translate(${dx}px,${dy}px) scale(.3)`, opacity: .15, offset: 1 }
  ];
  if (flyer.animate) {
    const flight = flyer.animate(keyframes, { duration: 900, easing: 'cubic-bezier(.2,.75,.3,1)', fill: 'forwards' });
    try { await flight.finished } catch { }
  } else await wait(900);
  flyer.remove();
  if (token !== animationToken) return false;
  placeNameInLiveDesk(item);
  const full = [...target.querySelectorAll('.live-seat')].every(seat => seat.textContent.trim());
  if (full) {
    target.classList.add('complete-celebration'); desks[item.physicalDesk].classList.add('desk-complete');
    await wait(1000);
    target.classList.remove('complete-celebration'); desks[item.physicalDesk].classList.remove('desk-complete');
  }
  return token === animationToken;
}
$('start-random').onclick = () => {
  if (selectedDesks.length * 2 < importedNames.length) {
    const missing = importedNames.length - selectedDesks.length * 2;
    $('random-error').textContent = `Thiếu ${missing} chỗ ngồi. Hãy chọn thêm ${Math.ceil(missing / 2)} bàn.`;
    return;
  }
  if (!selectedDesks.length) { $('random-error').textContent = 'Hãy chọn ít nhất một bàn.'; return; }
  $('random-close').disabled = true;
  $('export-random').disabled = true;
  clearArrangeHistory();
  randomAssignments = makeAssignments();
  seats.forEach(seat => seat.value = ''); updateCount();
  prepareLiveLayout();
  $('random-select-view').hidden = true; $('random-run-view').hidden = false;
  $('assignment-list').replaceChildren(); $('finish-random').hidden = true; $('skip-animation').hidden = false;
  $('random-progress').style.width = '0%'; $('progress-text').textContent = '';
  runAnimation();
};
function wait(ms) { return new Promise(resolve => setTimeout(resolve, ms)) }
async function animateDeskNumber(finalNumber, token) {
  const number = $('draw-number'), card = document.querySelector('.draw-card');
  number.classList.remove('landed'); number.classList.add('rolling'); card.classList.add('suspense');
  document.querySelector('.draw-label').textContent = 'ĐANG CHỌN BÀN…';
  const started = performance.now(), duration = 1900;
  while (performance.now() - started < duration) {
    if (token !== animationToken) return false;
    const progress = (performance.now() - started) / duration;
    number.textContent = String(1 + Math.floor(Math.random() * selectedDesks.length));
    await wait(55 + Math.pow(progress, 2) * 190);
  }
  if (token !== animationToken) return false;
  number.classList.remove('rolling'); number.textContent = String(finalNumber);
  void number.offsetWidth;
  number.classList.add('landed'); card.classList.remove('suspense'); document.querySelector('.draw-label').textContent = 'KẾT QUẢ';
  await wait(700);
  return token === animationToken;
}
async function runAnimation() {
  const token = ++animationToken;
  for (let i = 0; i < randomAssignments.length; i++) {
    if (token !== animationToken) return;
    const item = randomAssignments[i];
    const name = $('draw-name'); name.classList.remove('revealing'); name.textContent = item.name; void name.offsetWidth; name.classList.add('revealing');
    if (!await animateDeskNumber(item.deskNumber, token)) return;
    if (!await flyNameToDesk(item, token)) return;
    addAssignmentRow(item, i); applyOneAssignment(item);
    const done = i + 1;
    $('random-progress').style.width = `${done / randomAssignments.length * 100}%`;
    $('progress-text').textContent = `${done} / ${randomAssignments.length} học sinh`;
    await wait(450);
  }
  if (token === animationToken) finishAnimation();
}
function addAssignmentRow(item, i) {
  const row = document.createElement('div');
  const left = document.createElement('span'), number = document.createElement('b'), name = document.createTextNode(` ${item.name}`);
  number.textContent = `${i + 1}.`; left.append(number, name);
  const right = document.createElement('strong'); right.textContent = `Bàn ${item.deskNumber}`;
  row.append(left, right); $('assignment-list').append(row); $('assignment-list').scrollTop = $('assignment-list').scrollHeight;
}
function applyOneAssignment(item) { seats[item.seat].value = item.name; updateCount() }
function applyAllAssignments() {
  seats.forEach(seat => seat.value = '');
  prepareLiveLayout(); randomAssignments.forEach(placeNameInLiveDesk);
  randomAssignments.forEach(applyOneAssignment);
  $('assignment-list').replaceChildren(); randomAssignments.forEach(addAssignmentRow);
}
function finishAnimation() {
  $('random-close').disabled = false;
  refreshDeskLabels();
  $('skip-animation').hidden = true; $('finish-random').hidden = false;
  $('draw-name').textContent = 'Đã random xong!'; $('draw-number').textContent = '✓'; $('draw-number').classList.remove('rolling', 'landed');
  document.querySelector('.draw-card').classList.remove('suspense'); document.querySelector('.draw-label').textContent = 'HOÀN TẤT';
  $('random-progress').style.width = '100%'; $('progress-text').textContent = `Đã xếp ${randomAssignments.length} học sinh`;
  $('export-random').disabled = false; $('status').textContent = `Đã random chỗ ngồi cho ${randomAssignments.length} học sinh.`;
}
$('skip-animation').onclick = () => { animationToken++; applyAllAssignments(); finishAnimation() };
$('finish-random').onclick = closeRandom;

$('export-random').onclick = () => {
  if (!randomAssignments.length) return;
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; return; }
  const rows = [['STT', 'Họ và tên', 'Bàn', 'Vị trí trên sơ đồ']];
  randomAssignments.forEach((item, i) => rows.push([i + 1, item.name, item.deskNumber, physicalDeskName(item.physicalDesk)]));
  const book = XLSX.utils.book_new(), sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [{ wch: 7 }, { wch: 36 }, { wch: 10 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(book, sheet, 'Kết quả random');
  XLSX.writeFile(book, 'ket-qua-random-' + ($('className').value.trim() || 'lop') + '.xlsx');
};

// Tải Word từ mẫu gốc. Không gửi dữ liệu lên máy chủ.
$('word').onclick = async () => {
  const button = $('word'); button.disabled = true; button.textContent = 'Đang tạo Word…';
  try {
    if (!window.JSZip) throw Error('Thiếu file jszip.min.js.');
    const data = { teacher: $('teacher').value.trim(), className: $('className').value.trim(), year: $('year').value.trim(), names: seats.map(s => s.value) };
    const blob = await makeClassDocx(data);
    const url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = 'so-do-lop-' + (data.className || 'moi').replace(/[^\p{L}\p{N}_-]/gu, '-') + '.docx';
    document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 30000);
    $('status').textContent = 'Đã tạo file Word theo sơ đồ hiện tại.';
  } catch (error) { $('status').textContent = 'Không tạo được Word. Hãy giải nén đầy đủ các file rồi mở lại index.html.'; console.error(error) }
  finally { button.disabled = false; button.textContent = 'Tải Word (.docx)' }
};

updateHeader(); updateCount();
