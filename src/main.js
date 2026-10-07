import ExcelJS from "exceljs";
import JSZip from "jszip";
import "./style.css";

const asset = (name) => `${import.meta.env.BASE_URL}assets/${name}`;
let photoDataUrl = "";

const field = (id, label, type = "text", required = false, options = []) => {
  const control = options.length
    ? `<select id="${id}" ${required ? "required" : ""}><option value="">请选择</option>${options.map((item) => `<option>${item}</option>`).join("")}</select>`
    : `<input id="${id}" type="${type}" ${required ? "required" : ""} />`;
  return `<label class="field"><span>${label}</span>${control}</label>`;
};

document.querySelector("#app").innerHTML = `
  <main class="shell">
    <header class="hero">
      <p class="eyebrow">哈尔滨工业大学军乐团</p>
      <h1>团员档案生成器</h1>
      <p>填写一次，下载可直接导入 Obsidian 档案库的完整 ZIP 包。</p>
    </header>

    <section class="privacy" aria-label="隐私说明">
      <strong>本地处理</strong><span>填写内容与照片只在当前浏览器中使用，不会上传、保存到服务器或写入浏览器存储。关闭或刷新页面后即清空。</span>
    </section>

    <form id="profile-form">
      <section class="card">
        <div class="section-title"><span>01</span><div><h2>基本档案</h2><p>除老师外，其余项目均为必填。</p></div></div>
        <div class="form-grid">
          ${field("name", "姓名", "text", true)}
          ${field("studentId", "学号", "text", true)}
          ${field("part", "声部", "text", true)}
          ${field("joinDate", "入团时间", "date", true)}
          ${field("foundation", "是否零基础", "text", true, ["是", "否"])}
          ${field("teacher", "老师")}
          ${field("phone", "电话", "tel", true)}
          ${field("hometown", "家乡", "text", true)}
          ${field("qq", "QQ", "text", true)}
          ${field("wechat", "微信", "text", true)}
          ${field("birthday", "生日", "date", true)}
          ${field("college", "所在学院（部）", "text", true)}
          <label class="field photo-field"><span>证件照</span><input id="photo" type="file" accept="image/png,image/jpeg" required /><small>支持 PNG、JPG，导出时自动适配照片粘贴处。</small></label>
        </div>
        <div class="photo-preview" id="photo-preview" hidden><img alt="证件照预览" /><button type="button" id="remove-photo">移除照片</button></div>
      </section>

      <section class="actions">
        <div><h2>导出档案包</h2><p>下载一个 ZIP，内含 Markdown 档案、Word 档案表、Excel 台账和声部文本文件。</p></div>
        <div class="buttons"><button type="button" class="primary" id="download-package">下载 Obsidian 档案包</button></div>
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
    name: getValue("name"), studentId: getValue("studentId"), part: getValue("part"),
    joinDate: dateText(getValue("joinDate")), foundation: getValue("foundation"), teacher: getValue("teacher"),
    phone: getValue("phone"), hometown: getValue("hometown"), qq: getValue("qq"), wechat: getValue("wechat"),
    birthday: dateText(getValue("birthday")), college: getValue("college")
  };
}
function valid() {
  const form = document.querySelector("#profile-form");
  if (!form.reportValidity()) { setStatus("请完整填写除老师外的所有项目，并上传证件照。", true); return false; }
  return true;
}
function setStatus(message, error = false) {
  const target = document.querySelector("#status"); target.textContent = message; target.classList.toggle("error", error);
}
function filename(d, extension) { return `军乐团团员档案_${d.name || "未命名"}_${d.studentId || "档案"}.${extension}`; }
function safePathSegment(value, fallback) {
  return String(value || fallback).replace(/[\\/:*?"<>|]/g, "_").trim() || fallback;
}
function markdownText(value) {
  return String(value ?? "").replaceAll("|", "\\|").replaceAll("\n", " ");
}
function birthdayTags(value) {
  const match = String(value).match(/^\d{4}\.(\d{2})\.(\d{2})$/);
  return match ? `#${Number(match[1])}月 #${Number(match[2])}日` : "";
}

function imageExtension(dataUrl) { return dataUrl.startsWith("data:image/png") ? "png" : "jpeg"; }
async function createExcel(d) {
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet("人员汇总（含照片）", { views: [{ state: "frozen", ySplit: 1 }] });
    const headers = ["姓名", "学号", "声部", "入团时间", "是否零基础", "老师", "电话", "家乡", "QQ", "微信", "生日", "所在学院（部）", "人物照片"];
    sheet.addRow(headers);
    sheet.addRow([d.name, d.studentId, d.part, d.joinDate, d.foundation, d.teacher, d.phone, d.hometown, d.qq, d.wechat, d.birthday, d.college, ""]);
    sheet.columns = [12, 18, 12, 16, 15, 14, 18, 22, 16, 22, 16, 22, 15].map((width) => ({ width }));
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
      sheet.addImage(imageId, { tl: { col: 12.12, row: 1.07 }, ext: { width: 76, height: 92 } });
    } else {
      const photoCell = sheet.getCell("M2"); photoCell.value = "照片粘贴处"; photoCell.alignment = { vertical: "middle", horizontal: "center" };
    }
    return book.xlsx.writeBuffer();
}

function xmlEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}
async function photoAsPngBlob(dataUrl) {
  const image = await new Promise((resolve, reject) => {
    const source = new Image(); source.onload = () => resolve(source); source.onerror = reject; source.src = dataUrl;
  });
  const width = 240; const height = 320; const targetRatio = width / height; const sourceRatio = image.width / image.height;
  let sx = 0; let sy = 0; let sw = image.width; let sh = image.height;
  if (sourceRatio > targetRatio) { sw = image.height * targetRatio; sx = (image.width - sw) / 2; }
  else { sh = image.width / targetRatio; sy = (image.height - sh) / 2; }
  const canvas = document.createElement("canvas"); canvas.width = width; canvas.height = height;
  canvas.getContext("2d").drawImage(image, sx, sy, sw, sh, 0, 0, width, height);
  return new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("无法转换照片")), "image/png"));
}

async function createWord(d) {
    const [template, config] = await Promise.all([
      fetch(asset("member-profile-word-template-v2.docx"), { cache: "no-store" }).then((r) => r.arrayBuffer()),
      fetch(asset("member-profile-word-template-v2.json"), { cache: "no-store" }).then((r) => r.json())
    ]);
    const zip = await JSZip.loadAsync(template);
    const documentXml = zip.file("word/document.xml");
    if (!documentXml) throw new Error("找不到 Word 模板正文。");
    let xml = await documentXml.async("string");
    xml = xml.replaceAll("{{number}}", "");
    Object.entries(d).forEach(([key, value]) => { xml = xml.replaceAll(`{{${key}}}`, xmlEscape(value)); });
    zip.file("word/document.xml", xml);
    if (photoDataUrl) zip.file(config.photoMediaPath, await photoAsPngBlob(photoDataUrl));
    return zip.generateAsync({ type: "blob", compression: "DEFLATE" });
}

function createObsidianNote(d) {
  return `# ${markdownText(d.name)}\n\n## 一、基本信息\n\n| 项目 | 内容 |\n| --- | --- |\n| 学号 | ${markdownText(d.studentId)} |\n| 声部 | #${markdownText(d.part)} |\n| 出生月份 | ${birthdayTags(d.birthday)} |\n| 入团时间 | ${markdownText(d.joinDate)} |\n| 是否零基础 | #${markdownText(d.foundation)} |\n| 老师 | ${markdownText(d.teacher)} |\n| 电话 | ${markdownText(d.phone)} |\n| 家乡 | #${markdownText(d.hometown)} |\n| QQ | ${markdownText(d.qq)} |\n| 微信 | ${markdownText(d.wechat)} |\n| 所在学院（部） | ${markdownText(d.college)} |\n\n## 二、在乐团何时何地受到何种奖励\n\n| 序号 | 时间 | 奖励名称 |\n| --- | --- | --- |\n\n## 三、团员排练考勤\n\n> [!NOTE] 考核标准\n> 出勤率60%及以上为及格、出勤率80%及以上为良好、出勤率90%及以上为优秀\n\n| 学期 | 出勤率 | 评级 |\n| --- | --- | --- |\n\n## 四、演出登记\n\n| 序号 | 演出主题 | 曲目名称（专场无需填写） |\n| --- | --- | --- |\n\n## 五、何时担任何种职务\n\n| 序号 | 任期 | 职务 |\n| --- | --- | --- |\n`;
}

document.querySelector("#download-package").addEventListener("click", async () => {
  if (!valid()) return;
  const button = document.querySelector("#download-package");
  button.disabled = true; setStatus("正在生成 Obsidian 档案包，请稍候…");
  try {
    const d = data();
    const memberName = safePathSegment(d.name, "未命名");
    const part = safePathSegment(d.part, "未分类声部");
    const markdownFilename = `${memberName}.md`;
    const partFilename = `${part}.txt`;
    const [word, excel] = await Promise.all([createWord(d), createExcel(d)]);
    const archive = new JSZip();
    archive.file(markdownFilename, createObsidianNote(d));
    archive.file(filename(d, "docx"), word);
    archive.file(filename(d, "xlsx"), excel);
    archive.file(partFilename, `声部：${d.part}\r\n`);
    const blob = await archive.generateAsync({ type: "blob", compression: "DEFLATE" });
    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = filename(d, "zip"); link.click(); URL.revokeObjectURL(link.href);
    setStatus("Obsidian 档案包已开始下载。请转交管理员，再将 Markdown 文件放入对应声部文件夹。");
  } catch (error) { console.error(error); setStatus("档案包生成失败，请刷新页面后重试。", true); }
  finally { button.disabled = false; }
});
