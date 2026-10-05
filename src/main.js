import { PDFDocument, rgb } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as XLSX from "xlsx";
import "./style.css";

const asset = (name) => `${import.meta.env.BASE_URL}assets/${name}`;
let photoDataUrl = "";

const field = (id, label, type = "text", required = false, options = []) => {
  const control = options.length
    ? `<select id="${id}" ${required ? "required" : ""}><option value="">请选择</option>${options.map((item) => `<option>${item}</option>`).join("")}</select>`
    : `<input id="${id}" type="${type}" ${required ? "required" : ""} />`;
  return `<label class="field"><span>${label}${required ? '<b aria-label="必填">*</b>' : ""}</span>${control}</label>`;
};

document.querySelector("#app").innerHTML = `
  <main class="shell">
    <header class="hero">
      <p class="eyebrow">哈尔滨工业大学军乐团</p>
      <h1>团员档案生成器</h1>
      <p>填写一次，即可在本机下载标准 PDF 档案表和 Excel 台账。</p>
    </header>

    <section class="privacy" aria-label="隐私说明">
      <strong>本地处理</strong><span>填写内容与照片只在当前浏览器中使用，不会上传或保存到服务器。</span>
    </section>

    <form id="profile-form">
      <section class="card">
        <div class="section-title"><span>01</span><div><h2>基本档案</h2><p>带 * 的项目会出现在档案表中。</p></div></div>
        <div class="form-grid">
          ${field("number", "编号")}
          ${field("name", "姓名", "text", true)}
          ${field("studentId", "学号", "text", true)}
          ${field("part", "声部", "text", true)}
          ${field("joinDate", "入团时间", "date")}
          ${field("foundation", "是否零基础", "text", false, ["是", "否"])}
          ${field("teacher", "老师")}
          ${field("phone", "电话", "tel")}
          ${field("hometown", "家乡")}
          ${field("qq", "QQ")}
          ${field("wechat", "微信")}
          <label class="field photo-field"><span>证件照</span><input id="photo" type="file" accept="image/png,image/jpeg" /><small>支持 PNG、JPG，导出时自动适配照片粘贴处。</small></label>
        </div>
        <div class="photo-preview" id="photo-preview" hidden><img alt="证件照预览" /><button type="button" id="remove-photo">移除照片</button></div>
      </section>

      ${recordSection("awards", "02", "获奖记录", "时间", "奖励名称", "最多显示 6 条，按时间顺序填写。")}
      ${attendanceSection()}
      ${recordSection("performances", "04", "演出登记", "时间", "地点", "演出主题", "曲目名称（专场无需填写）", "最多显示 9 条。")}

      <section class="actions">
        <div><h2>导出档案</h2><p>确认信息后分别下载 PDF 或 Excel。</p></div>
        <div class="buttons"><button type="button" class="secondary" id="download-xlsx">下载 Excel</button><button type="button" class="primary" id="download-pdf">下载 PDF</button></div>
      </section>
    </form>
    <p id="status" class="status" role="status" aria-live="polite"></p>
  </main>
`;

function recordSection(id, no, title, ...columns) {
  const help = columns.pop();
  const headings = columns;
  return `<section class="card records" data-records="${id}" data-limit="${id === "awards" ? 6 : 9}">
    <div class="section-title"><span>${no}</span><div><h2>${title}</h2><p>${help}</p></div></div>
    <div class="record-head" style="--columns:${headings.length}">${headings.map((name) => `<span>${name}</span>`).join("")}<span>操作</span></div>
    <div class="record-list"></div><button type="button" class="add-record">添加一条</button>
  </section>`;
}

function attendanceSection() {
  return `<section class="card records" data-records="attendance" data-limit="6">
    <div class="section-title"><span>03</span><div><h2>团员排练考勤</h2><p>最多显示 6 条。考勤标记与原表保持一致。</p></div></div>
    <div class="record-head" style="--columns:3"><span>学期</span><span>考勤标记</span><span>声部长签字</span><span>操作</span></div>
    <div class="record-list"></div><button type="button" class="add-record">添加一条</button>
  </section>`;
}

function addRecord(section) {
  const id = section.dataset.records;
  const list = section.querySelector(".record-list");
  if (list.children.length >= Number(section.dataset.limit)) return setStatus(`该部分最多填写 ${section.dataset.limit} 条记录。`, true);
  const row = document.createElement("div");
  row.className = "record-row";
  if (id === "awards") row.innerHTML = `<input type="date" aria-label="获奖时间"><input aria-label="奖励名称"><button type="button" class="remove-record">删除</button>`;
  if (id === "attendance") row.innerHTML = `<input aria-label="学期" placeholder="如：2026 春季"><select aria-label="考勤标记"><option value="">请选择</option><option>90%优</option><option>80%良</option><option>60%及格</option></select><input aria-label="声部长签字"><button type="button" class="remove-record">删除</button>`;
  if (id === "performances") row.innerHTML = `<input type="date" aria-label="演出时间"><input aria-label="地点"><input aria-label="演出主题"><input aria-label="曲目名称"><button type="button" class="remove-record">删除</button>`;
  list.append(row);
}

document.querySelectorAll("[data-records]").forEach((section) => addRecord(section));
document.addEventListener("click", (event) => {
  const section = event.target.closest("[data-records]");
  if (event.target.matches(".add-record")) addRecord(section);
  if (event.target.matches(".remove-record")) event.target.closest(".record-row").remove();
});

document.querySelector("#photo").addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!/image\/(png|jpeg)/.test(file.type)) return setStatus("请上传 PNG 或 JPG 格式的照片。", true);
  const reader = new FileReader();
  reader.onload = () => {
    photoDataUrl = String(reader.result);
    const preview = document.querySelector("#photo-preview");
    preview.querySelector("img").src = photoDataUrl;
    preview.hidden = false;
  };
  reader.readAsDataURL(file);
});
document.querySelector("#remove-photo").addEventListener("click", () => {
  photoDataUrl = "";
  document.querySelector("#photo").value = "";
  document.querySelector("#photo-preview").hidden = true;
});

function getValue(id) { return document.querySelector(`#${id}`).value.trim(); }
function dateText(value) { return value ? value.replaceAll("-", ".") : ""; }
function collectRows(id) {
  return [...document.querySelector(`[data-records="${id}"] .record-list`).children]
    .map((row) => [...row.querySelectorAll("input,select")].map((input) => input.value.trim()))
    .filter((values) => values.some(Boolean));
}
function data() {
  return {
    number: getValue("number"), name: getValue("name"), studentId: getValue("studentId"), part: getValue("part"),
    joinDate: dateText(getValue("joinDate")), foundation: getValue("foundation"), teacher: getValue("teacher"),
    phone: getValue("phone"), hometown: getValue("hometown"), qq: getValue("qq"), wechat: getValue("wechat"),
    awards: collectRows("awards").map(([date, name]) => [dateText(date), name]), attendance: collectRows("attendance"),
    performances: collectRows("performances").map(([date, place, theme, song]) => [dateText(date), place, theme, song])
  };
}
function valid() {
  const form = document.querySelector("#profile-form");
  if (!form.reportValidity()) { setStatus("请先填写姓名、学号和声部。", true); return false; }
  return true;
}
function setStatus(message, error = false) {
  const target = document.querySelector("#status"); target.textContent = message; target.classList.toggle("error", error);
}
function filename(d, extension) { return `军乐团团员档案_${d.name || "未命名"}_${d.studentId || "档案"}.${extension}`; }

function worksheet(rows, widths) {
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = widths.map((width) => ({ wch: width }));
  return sheet;
}
document.querySelector("#download-xlsx").addEventListener("click", () => {
  if (!valid()) return;
  const d = data();
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, worksheet([["哈尔滨工业大学军乐团团员档案"], [], ["编号", d.number], ["姓名", d.name], ["学号", d.studentId], ["声部", d.part], ["入团时间", d.joinDate], ["是否零基础", d.foundation], ["老师", d.teacher], ["电话", d.phone], ["家乡", d.hometown], ["QQ", d.qq], ["微信", d.wechat]], [18, 32]), "基本档案");
  XLSX.utils.book_append_sheet(book, worksheet([["序号", "时间", "奖励名称"], ...d.awards.map((row, index) => [index + 1, ...row])], [10, 18, 48]), "获奖记录");
  XLSX.utils.book_append_sheet(book, worksheet([["学期", "考勤标记", "声部长签字"], ...d.attendance], [20, 20, 28]), "排练考勤");
  XLSX.utils.book_append_sheet(book, worksheet([["时间", "地点", "演出主题", "曲目名称（专场无需填写）"], ...d.performances], [18, 24, 36, 42]), "演出登记");
  XLSX.writeFile(book, filename(d, "xlsx"));
  setStatus("Excel 已开始下载。");
});

function fit(text, font, size, width) {
  let value = String(text || "");
  while (value && font.widthOfTextAtSize(value, size) > width) value = `${value.slice(0, -2)}…`;
  return value;
}
function draw(page, font, value, x, y, width, size = 8.4) {
  const text = fit(value, font, size, width);
  page.drawText(text, { x, y, size, font, color: rgb(0.12, 0.12, 0.12) });
}
async function imageBytes(dataUrl) { return Uint8Array.from(atob(dataUrl.split(",")[1]), (c) => c.charCodeAt(0)); }

document.querySelector("#download-pdf").addEventListener("click", async () => {
  if (!valid()) return;
  const button = document.querySelector("#download-pdf");
  button.disabled = true; setStatus("正在生成 PDF，请稍候…");
  try {
    const d = data();
    const [template, fontBytes] = await Promise.all([
      fetch(asset("member-profile-template.pdf")).then((r) => r.arrayBuffer()),
      fetch(asset("NotoSansCJKsc-Regular.otf")).then((r) => r.arrayBuffer())
    ]);
    const pdf = await PDFDocument.load(template); pdf.registerFontkit(fontkit);
    const font = await pdf.embedFont(fontBytes, { subset: true }); const page = pdf.getPage(0);
    draw(page, font, d.number, 474, 704, 62, 9);
    [[d.name, 140, 686, 68], [d.studentId, 320, 686, 60], [d.part, 140, 664, 68], [d.joinDate, 320, 664, 60], [d.foundation, 140, 642, 68], [d.teacher, 320, 642, 60], [d.phone, 140, 620, 68], [d.hometown, 320, 620, 60], [d.qq, 140, 598, 68], [d.wechat, 320, 598, 60]].forEach(([value, x, y, width]) => draw(page, font, value, x, y, width));
    if (photoDataUrl) {
      const bytes = await imageBytes(photoDataUrl); const photo = photoDataUrl.startsWith("data:image/png") ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
      page.drawImage(photo, { x: 386, y: 610, width: 66, height: 89 });
    }
    d.awards.slice(0, 6).forEach(([date, name], i) => { const y = 569 - i * 18; draw(page, font, i + 1, 87, y, 18); draw(page, font, date, 121, y, 47); draw(page, font, name, 300, y, 148); });
    d.attendance.slice(0, 6).forEach(([term, mark, signer], i) => { const y = 447 - i * 18; draw(page, font, term, 92, y, 67); draw(page, font, mark, 215, y, 151); draw(page, font, signer, 397, y, 44); });
    d.performances.slice(0, 9).forEach(([date, place, theme, song], i) => { const y = 310 - i * 18; draw(page, font, date, 88, y, 55); draw(page, font, place, 152, y, 60); draw(page, font, theme, 226, y, 102); draw(page, font, song, 338, y, 108); });
    const blob = new Blob([await pdf.save()], { type: "application/pdf" });
    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = filename(d, "pdf"); link.click(); URL.revokeObjectURL(link.href);
    setStatus("PDF 已开始下载。");
  } catch (error) { console.error(error); setStatus("PDF 生成失败，请刷新页面后重试。", true); }
  finally { button.disabled = false; }
});
