import { PDFDocument, rgb } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import ExcelJS from "exceljs";
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

      <section class="actions">
        <div><h2>导出档案</h2><p>确认信息后分别下载 PDF 或 Excel。</p></div>
        <div class="buttons"><button type="button" class="secondary" id="download-xlsx">下载 Excel</button><button type="button" class="primary" id="download-pdf">下载 PDF</button></div>
      </section>
    </form>
    <p id="status" class="status" role="status" aria-live="polite"></p>
  </main>
`;

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
function data() {
  return {
    number: getValue("number"), name: getValue("name"), studentId: getValue("studentId"), part: getValue("part"),
    joinDate: dateText(getValue("joinDate")), foundation: getValue("foundation"), teacher: getValue("teacher"),
    phone: getValue("phone"), hometown: getValue("hometown"), qq: getValue("qq"), wechat: getValue("wechat")
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

function imageExtension(dataUrl) { return dataUrl.startsWith("data:image/png") ? "png" : "jpeg"; }
document.querySelector("#download-xlsx").addEventListener("click", async () => {
  if (!valid()) return;
  const button = document.querySelector("#download-xlsx");
  button.disabled = true; setStatus("正在生成汇总 Excel，请稍候…");
  try {
    const d = data(); const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet("人员汇总（含照片）", { views: [{ state: "frozen", ySplit: 1 }] });
    const headers = ["姓名", "学号", "声部", "入团时间", "是否零基础", "老师", "电话", "家乡", "QQ", "微信", "人物照片"];
    sheet.addRow(headers);
    sheet.addRow([d.name, d.studentId, d.part, d.joinDate, d.foundation, d.teacher, d.phone, d.hometown, d.qq, d.wechat, ""]);
    sheet.columns = [12, 18, 12, 16, 15, 14, 18, 22, 16, 22, 15].map((width) => ({ width }));
    sheet.getRow(1).height = 24; sheet.getRow(2).height = 96;
    sheet.getRow(1).eachCell((cell) => {
      cell.font = { name: "宋体", bold: true, size: 11, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2F6B4F" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { top: { style: "thin", color: { argb: "FFB7C9BD" } }, left: { style: "thin", color: { argb: "FFB7C9BD" } }, bottom: { style: "thin", color: { argb: "FFB7C9BD" } }, right: { style: "thin", color: { argb: "FFB7C9BD" } } };
    });
    sheet.getRow(2).eachCell((cell) => {
      cell.font = { name: "宋体", size: 11 }; cell.alignment = { vertical: "top", wrapText: true };
      cell.border = { top: { style: "thin", color: { argb: "FFD6E0D9" } }, left: { style: "thin", color: { argb: "FFD6E0D9" } }, bottom: { style: "thin", color: { argb: "FFD6E0D9" } }, right: { style: "thin", color: { argb: "FFD6E0D9" } } };
    });
    if (photoDataUrl) {
      const imageId = book.addImage({ base64: photoDataUrl, extension: imageExtension(photoDataUrl) });
      sheet.addImage(imageId, { tl: { col: 10.12, row: 1.07 }, ext: { width: 76, height: 92 } });
    } else {
      const photoCell = sheet.getCell("K2"); photoCell.value = "照片粘贴处"; photoCell.alignment = { vertical: "middle", horizontal: "center" };
    }
    const buffer = await book.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = filename(d, "xlsx"); link.click(); URL.revokeObjectURL(link.href);
    setStatus("人员汇总 Excel 已开始下载，可直接复制该行到汇总表末尾。");
  } catch (error) { console.error(error); setStatus("Excel 生成失败，请刷新页面后重试。", true); }
  finally { button.disabled = false; }
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
    [[d.name, 140, 676, 68], [d.studentId, 320, 676, 60], [d.part, 140, 654, 68], [d.joinDate, 320, 654, 60], [d.foundation, 140, 632, 68], [d.teacher, 320, 632, 60], [d.phone, 140, 610, 68], [d.hometown, 320, 610, 60], [d.qq, 140, 588, 68], [d.wechat, 320, 588, 60]].forEach(([value, x, y, width]) => draw(page, font, value, x, y, width));
    if (photoDataUrl) {
      const bytes = await imageBytes(photoDataUrl); const photo = photoDataUrl.startsWith("data:image/png") ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
      page.drawImage(photo, { x: 386, y: 598, width: 66, height: 89 });
    }
    const blob = new Blob([await pdf.save()], { type: "application/pdf" });
    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = filename(d, "pdf"); link.click(); URL.revokeObjectURL(link.href);
    setStatus("PDF 已开始下载。");
  } catch (error) { console.error(error); setStatus("PDF 生成失败，请刷新页面后重试。", true); }
  finally { button.disabled = false; }
});
