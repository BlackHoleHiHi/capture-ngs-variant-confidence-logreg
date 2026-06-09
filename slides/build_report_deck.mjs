import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspace = path.resolve("D:/learnDP/outputs/manual-hust-paper-report/presentations/variant-confidence-report");
const helperPath = path.join(workspace, "tools", "artifact_tool_utils.mjs");
const {
  createSlideContext,
  ensureArtifactToolWorkspace,
  importArtifactTool,
  saveBlobToFile,
} = await import(pathToFileURL(helperPath).href);

const starterPptx = path.join(workspace, "output", "template-starter.pptx");
const outputPptx = path.join(workspace, "output", "variant-confidence-hust-report.pptx");
const previewDir = path.join(workspace, "preview", "final");
const layoutDir = path.join(workspace, "layout", "final");
const manifestPath = path.join(workspace, "output", "variant-confidence-hust-report.manifest.json");

const RED = "#A30D12";
const RED_DARK = "#7D0710";
const GOLD = "#D89A2B";
const INK = "#242424";
const MUTED = "#666666";
const BLUE = "#24547A";
const GREEN = "#2D7A45";
const AMBER = "#8A5A00";
const LIGHT_RED = "#F7E7E8";
const LIGHT_BLUE = "#E9F1F7";
const LIGHT_GREEN = "#E8F4EB";
const LIGHT_GOLD = "#FFF3D8";
const BORDER = "#D7D7D7";

await ensureArtifactToolWorkspace(workspace);
const artifact = await importArtifactTool(workspace);
const { FileBlob, PresentationFile } = artifact;
const presentation = await PresentationFile.importPptx(await FileBlob.load(starterPptx));
const slideSize = { width: 960, height: 720 };
const ctx = createSlideContext(artifact, {
  slideSize,
  workspaceDir: workspace,
  assetDir: path.join(workspace, "assets"),
  outputDir: path.join(workspace, "output"),
  titleFont: "Lato",
  bodyFont: "Calibri",
  monoFont: "Consolas",
});

function slideAt(index) {
  return presentation.slides.getItem(index);
}

function shapeByName(slide, namePart) {
  return slide.shapes.items.find((shape) => String(shape.name || "").includes(namePart));
}

function deleteWhere(slide, predicate) {
  for (const shape of [...slide.shapes.items]) {
    if (predicate(shape)) shape.delete();
  }
}

function clearTemplatePlaceholders(slide) {
  deleteWhere(slide, (shape) => {
    const name = String(shape.name || "");
    if (name.includes("Slide Number")) return true;
    if (name.includes("Date Placeholder")) return true;
    if (name.includes("Footer Placeholder")) return true;
    return false;
  });
}

function deleteContentPlaceholders(slide) {
  deleteWhere(slide, (shape) => {
    const name = String(shape.name || "");
    if (name.includes("Title")) return false;
    if (name.includes("Placeholder")) return true;
    return false;
  });
}

function setExistingText(shape, text, options = {}) {
  if (!shape) return;
  shape.text = text;
  shape.text.fontSize = options.fontSize ?? 24;
  shape.text.color = options.color ?? INK;
  shape.text.bold = Boolean(options.bold ?? false);
  shape.text.typeface = options.typeface ?? "Calibri";
  shape.text.alignment = options.align ?? "left";
  shape.text.verticalAlignment = options.valign ?? "top";
  shape.text.insets = options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
}

function setTitle(slide, title, fontSize = 30) {
  const titleShape = slide.shapes.items.find((shape) => String(shape.name || "").includes("Title"));
  setExistingText(titleShape, title, {
    fontSize,
    color: "#FFFFFF",
    bold: true,
    typeface: "Lato",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  if (titleShape?.bringToFront) titleShape.bringToFront();
}

function addText(slide, text, x, y, w, h, options = {}) {
  return ctx.addText(slide, {
    text,
    x,
    y,
    w,
    h,
    fontSize: options.fontSize ?? 19,
    color: options.color ?? INK,
    bold: options.bold ?? false,
    typeface: options.typeface ?? "Calibri",
    align: options.align ?? "left",
    valign: options.valign ?? "top",
    fill: options.fill ?? "#00000000",
    line: options.line ?? { style: "solid", fill: "#00000000", width: 0 },
    insets: options.insets ?? { left: 8, right: 8, top: 6, bottom: 6 },
    name: options.name,
  });
}

function addBox(slide, x, y, w, h, options = {}) {
  return ctx.addShape(slide, {
    x,
    y,
    w,
    h,
    fill: options.fill ?? "#FFFFFF",
    line: { style: "solid", fill: options.stroke ?? BORDER, width: options.strokeWidth ?? 1 },
    name: options.name,
  });
}

function addFullBackground(slide, fill = "#FFFFFF") {
  const bg = addBox(slide, 0, 0, 960, 720, {
    fill,
    stroke: fill,
    strokeWidth: 0,
    name: "explicit white background",
  });
  if (bg.sendToBack) bg.sendToBack();
  return bg;
}

function addContentChrome(slide) {
  addFullBackground(slide, "#FFFFFF");
  addBox(slide, 0, 0, 960, 52, { fill: RED, stroke: RED, strokeWidth: 0, name: "HUST red title bar" });
  addBox(slide, 0, 64, 960, 12, { fill: GOLD, stroke: GOLD, strokeWidth: 0, name: "HUST gold divider" });
  addText(slide, "HUST", 28, 682, 60, 18, {
    fontSize: 13,
    color: RED,
    bold: true,
    typeface: "Lato",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addBox(slide, 92, 690, 820, 1.2, { fill: "#D6A1A5", stroke: "#D6A1A5", strokeWidth: 0 });
}

function addCoverChrome(slide) {
  addFullBackground(slide, "#FFFFFF");
  addBox(slide, 850, 0, 110, 720, { fill: "#F7E7E8", stroke: "#F7E7E8", strokeWidth: 0 });
  addBox(slide, 908, 0, 52, 720, { fill: "#FFFFFF", stroke: "#FFFFFF", strokeWidth: 0 });
  addBox(slide, 0, 0, 960, 8, { fill: RED, stroke: RED, strokeWidth: 0 });
}

function addClosingChrome(slide) {
  addFullBackground(slide, "#FFFFFF");
  addBox(slide, 0, 0, 320, 720, { fill: RED, stroke: RED, strokeWidth: 0 });
  addText(slide, "HUST", 78, 318, 170, 48, {
    fontSize: 34,
    color: "#FFFFFF",
    bold: true,
    typeface: "Lato",
    align: "center",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addBox(slide, 320, 0, 12, 720, { fill: GOLD, stroke: GOLD, strokeWidth: 0 });
}

function addCard(slide, title, body, x, y, w, h, options = {}) {
  addBox(slide, x, y, w, h, { fill: options.fill ?? "#FFFFFF", stroke: options.stroke ?? BORDER });
  addText(slide, title, x + 12, y + 10, w - 24, 28, {
    fontSize: options.titleSize ?? 17,
    color: options.titleColor ?? RED,
    bold: true,
    typeface: "Lato",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addText(slide, body, x + 12, y + 42, w - 24, h - 50, {
    fontSize: options.bodySize ?? 18,
    color: options.bodyColor ?? INK,
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function addMetric(slide, label, value, x, y, w, h, options = {}) {
  addBox(slide, x, y, w, h, { fill: options.fill ?? LIGHT_BLUE, stroke: options.stroke ?? "#B8CCD9" });
  addText(slide, value, x + 8, y + 8, w - 16, 36, {
    fontSize: options.valueSize ?? 25,
    color: options.valueColor ?? BLUE,
    bold: true,
    typeface: "Lato",
    align: "center",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addText(slide, label, x + 8, y + 47, w - 16, Math.max(12, h - 62), {
    fontSize: options.labelSize ?? 14,
    color: options.labelColor ?? MUTED,
    align: "center",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function addSmallSource(slide, text = "Nguồn: van den Akker et al., BMC Genomics 2018; PMID 29665779; Supplementary Table S2") {
  addText(slide, text, 36, 660, 850, 20, {
    fontSize: 10.5,
    color: "#777777",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function addPipelineStep(slide, number, title, body, x, y, w, fill) {
  addBox(slide, x, y, w, 92, { fill, stroke: BORDER });
  addText(slide, String(number), x + 10, y + 14, 34, 34, {
    fontSize: 20,
    color: "#FFFFFF",
    bold: true,
    align: "center",
    fill: RED,
    insets: { left: 0, right: 0, top: 3, bottom: 0 },
  });
  addText(slide, title, x + 52, y + 10, w - 62, 24, {
    fontSize: 17,
    color: RED_DARK,
    bold: true,
    typeface: "Lato",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addText(slide, body, x + 52, y + 38, w - 62, 38, {
    fontSize: 12.8,
    color: INK,
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function addArrow(slide, x, y) {
  addText(slide, "->", x, y, 34, 28, {
    fontSize: 22,
    color: GOLD,
    bold: true,
    align: "center",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function addTable(slide, rows, x, y, widths, rowH, options = {}) {
  const headerFill = options.headerFill ?? RED;
  const fontSize = options.fontSize ?? 16;
  const totalW = widths.reduce((a, b) => a + b, 0);
  for (let r = 0; r < rows.length; r += 1) {
    const fill = r === 0 ? headerFill : r % 2 === 0 ? "#FAFAFA" : "#FFFFFF";
    addBox(slide, x, y + r * rowH, totalW, rowH, { fill, stroke: BORDER });
    let cx = x;
    for (let c = 0; c < rows[r].length; c += 1) {
      addText(slide, rows[r][c], cx + 6, y + r * rowH + 7, widths[c] - 12, rowH - 20, {
        fontSize,
        color: r === 0 ? "#FFFFFF" : INK,
        bold: r === 0 || (c === 0 && options.boldFirstCol),
        align: c === 0 ? "left" : "center",
        insets: { left: 0, right: 0, top: 0, bottom: 0 },
        valign: "middle",
      });
      cx += widths[c];
    }
  }
}

for (const slide of presentation.slides.items) {
  clearTemplatePlaceholders(slide);
}

// Slide 1: opening title
{
  const slide = slideAt(0);
  addCoverChrome(slide);
  const title = shapeByName(slide, "Title");
  setExistingText(title, "A machine learning model to determine the accuracy of variant calls in capture based next generation sequencing", {
    fontSize: 22,
    color: RED,
    bold: true,
    typeface: "Lato",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  addText(
    slide,
    "Logistic Regression cho đánh giá độ tin cậy variant call\nBMC Genomics 2018 | Dữ liệu: Supplementary Table S2 | Mục tiêu: mô hình bảo thủ, FP = 0",
    44,
    220,
    780,
    78,
    { fontSize: 17, color: INK, insets: { left: 0, right: 0, top: 0, bottom: 0 } },
  );
  addMetric(slide, "variants trong bài báo", "7,179", 44, 338, 178, 82, { fill: LIGHT_RED, valueColor: RED });
  addMetric(slide, "features sử dụng", "14", 246, 338, 178, 82, { fill: LIGHT_GOLD, valueColor: AMBER });
  addMetric(slide, "AUC chạy thử test", "0.999831", 448, 338, 214, 82, { fill: LIGHT_BLUE, valueColor: BLUE });
  addMetric(slide, "False Positive test", "0", 686, 338, 138, 82, { fill: LIGHT_GREEN, valueColor: GREEN });
  addSmallSource(slide);
}

// Slide 2: problem and objective
{
  const slide = slideAt(1);
  addContentChrome(slide);
  setTitle(slide, "Bài toán và yêu cầu mô hình", 29);
  deleteContentPlaceholders(slide);
  addCard(
    slide,
    "Vấn đề trong clinical NGS",
    "- NGS capture-based tạo nhiều variant call, nhưng chất lượng không đồng nhất.\n- Một nhóm nhỏ variant khó gọi có thể là artifact do depth, mapping, strand bias hoặc vùng lặp.\n- Báo cáo nhầm variant giả có rủi ro cao trong bối cảnh lâm sàng.",
    42,
    142,
    400,
    302,
    { fill: "#FFFFFF", titleColor: RED, bodySize: 18 },
  );
  addCard(
    slide,
    "Mục tiêu của bài báo",
    "- Học từ các variant đã được xác nhận bằng Sanger.\n- Gán High confidence cho variant có thể tin cậy bằng NGS.\n- Giữ Low confidence cho ca cần kiểm tra bổ sung.\n- Ưu tiên zero false-positive hơn tối đa recall.",
    500,
    142,
    388,
    302,
    { fill: "#FFFFFF", titleColor: RED, bodySize: 18 },
  );
  addBox(slide, 100, 484, 760, 70, { fill: LIGHT_RED, stroke: "#E6B7BA" });
  addText(slide, "Nguyên tắc báo cáo: High confidence chỉ khi xác suất >= ngưỡng bảo thủ và FP trên Dev bằng 0.", 122, 501, 716, 38, {
    fontSize: 18,
    color: RED_DARK,
    bold: true,
    align: "center",
  });
}

// Slide 3: data and features
{
  const slide = slideAt(2);
  addContentChrome(slide);
  setTitle(slide, "Dữ liệu và 14 feature sử dụng", 29);
  deleteContentPlaceholders(slide);
  addMetric(slide, "Train / Dev / Test", "70 / 15 / 15", 42, 92, 230, 72, { fill: LIGHT_BLUE, valueSize: 23 });
  addMetric(slide, "Target", "Present = 1", 298, 92, 190, 72, { fill: LIGHT_GREEN, valueSize: 23, valueColor: GREEN });
  addMetric(slide, "Negative class", "Not present = 0", 514, 92, 220, 72, { fill: LIGHT_RED, valueSize: 23, valueColor: RED });
  addMetric(slide, "Source", "Table S2", 760, 92, 150, 72, { fill: LIGHT_GOLD, valueSize: 22, valueColor: AMBER });
  addCard(slide, "Read support", "DP, AD, AF\nAF = AD / DP", 58, 206, 190, 132, { fill: "#FFFFFF", bodySize: 19 });
  addCard(slide, "Call quality", "QUAL, QD, MQ, GQ, FS\nQD = QUAL / DP", 282, 206, 250, 132, { fill: "#FFFFFF", bodySize: 18 });
  addCard(slide, "Sequence context", "gc_5, gc_20, gc_50\nHPL-D, HPL-L", 566, 206, 250, 132, { fill: "#FFFFFF", bodySize: 18 });
  addCard(slide, "Other signal", "WHR\nweighted homopolymer rate", 218, 382, 250, 126, { fill: "#FFFFFF", bodySize: 18 });
  addCard(slide, "Preprocessing", "Loại dòng metadata/rỗng -> encode y -> numeric coercion -> StandardScaler trên X.", 500, 382, 318, 126, {
    fill: "#FFFFFF",
    bodySize: 17,
  });
  addSmallSource(slide);
}

// Slide 4: pipeline
{
  const slide = slideAt(3);
  addContentChrome(slide);
  setTitle(slide, "Pipeline tái lập từ dữ liệu của bài báo", 29);
  deleteContentPlaceholders(slide);
  const xs = [44, 352, 660];
  addPipelineStep(slide, 1, "Input Excel", "12864_2018_4659_MOESM2_ESM.xlsx\nvariant-level Table S2", xs[0], 130, 245, LIGHT_BLUE);
  addArrow(slide, 300, 155);
  addPipelineStep(slide, 2, "Clean + encode", "Bỏ metadata/rỗng, Present=1,\nNot present=0, ép numeric", xs[1], 130, 245, LIGHT_GREEN);
  addArrow(slide, 608, 155);
  addPipelineStep(slide, 3, "Feature set", "14 biến: DP, AD, AF, GC,\nMQ, GQ, WHR, HPL, QUAL, QD, FS", xs[2], 130, 245, LIGHT_GOLD);
  addPipelineStep(slide, 4, "Split", "Train 70%, Dev 15%, Test 15%\nstratified theo target", xs[0], 320, 245, "#FFFFFF");
  addArrow(slide, 300, 345);
  addPipelineStep(slide, 5, "Model", "Pipeline(StandardScaler,\nLogisticRegression)", xs[1], 320, 245, "#FFFFFF");
  addArrow(slide, 608, 345);
  addPipelineStep(slide, 6, "Threshold", "Tìm ngưỡng thấp nhất trên Dev\nsao cho FP = 0", xs[2], 320, 245, LIGHT_RED);
  addBox(slide, 148, 548, 664, 52, { fill: "#F8F8F8", stroke: BORDER });
  addText(slide, "Test set chỉ dùng sau khi đã chốt model và threshold để tránh rò rỉ đánh giá.", 170, 562, 620, 26, {
    fontSize: 17,
    color: RED_DARK,
    bold: true,
    align: "center",
  });
}

// Slide 5: conservative threshold
{
  const slide = slideAt(4);
  addContentChrome(slide);
  setTitle(slide, "Chọn ngưỡng bảo thủ trên Development set", 29);
  deleteContentPlaceholders(slide);
  addCard(
    slide,
    "Hàm tìm threshold",
    "Duyệt các xác suất dự đoán trên Dev.\nChọn threshold thấp nhất sao cho:\nFP = count(y_dev=0 và p >= threshold) = 0.\n\nCách triển khai ổn định:\nthreshold = next probability lớn hơn max(p của negative trên Dev).",
    54,
    138,
    500,
    262,
    { fill: "#FFFFFF", titleColor: RED, bodySize: 18 },
  );
  addMetric(slide, "conservative_threshold", "0.4730547338", 590, 138, 260, 92, {
    fill: LIGHT_RED,
    valueSize: 24,
    valueColor: RED,
  });
  addMetric(slide, "Dev False Positive", "0", 590, 260, 120, 86, { fill: LIGHT_GREEN, valueColor: GREEN });
  addMetric(slide, "Dev High confidence", "995", 730, 260, 120, 86, { fill: LIGHT_BLUE, valueColor: BLUE });
  addBox(slide, 120, 455, 720, 74, { fill: LIGHT_GOLD, stroke: "#E4C989" });
  addText(slide, "Điểm quan trọng: threshold được tối ưu trên Dev, không dùng Test. Điều này giữ test set là đánh giá khách quan.", 148, 472, 664, 42, {
    fontSize: 17,
    color: AMBER,
    bold: true,
    align: "center",
  });
}

// Slide 6: test evaluation
{
  const slide = slideAt(5);
  addContentChrome(slide);
  setTitle(slide, "Đánh giá trên Test set", 29);
  deleteContentPlaceholders(slide);
  addMetric(slide, "ROC AUC", "0.999831", 62, 120, 170, 80, { fill: LIGHT_BLUE, valueColor: BLUE });
  addMetric(slide, "High confidence", "996 / 1077", 252, 120, 190, 80, { fill: LIGHT_GOLD, valueColor: AMBER });
  addMetric(slide, "False Positive", "0", 462, 120, 150, 80, { fill: LIGHT_GREEN, valueColor: GREEN });
  addMetric(slide, "False Negative", "4", 632, 120, 150, 80, { fill: LIGHT_RED, valueColor: RED });
  addText(slide, "Confusion Matrix [[TN, FP], [FN, TP]]", 92, 248, 360, 26, {
    fontSize: 18,
    bold: true,
    color: RED_DARK,
    typeface: "Lato",
  });
  addTable(
    slide,
    [
      ["", "Pred Low", "Pred High"],
      ["Actual Not present", "77", "0"],
      ["Actual Present", "4", "996"],
    ],
    92,
    292,
    [170, 140, 140],
    54,
    { fontSize: 17, boldFirstCol: true },
  );
  addCard(
    slide,
    "Diễn giải",
    "- Không có variant giả lọt vào High confidence trên Test.\n- 4 variant thật bị đẩy xuống Low confidence: chấp nhận được khi ưu tiên bảo thủ.\n- Tỷ lệ High confidence: 92.48%.",
    565,
    265,
    300,
    188,
    { fill: "#FFFFFF", bodySize: 17 },
  );
  addSmallSource(slide, "Kết quả local từ script D:\\A\\train_conservative_threshold.py");
}

// Slide 7: comparison with paper
{
  const slide = slideAt(6);
  addContentChrome(slide);
  setTitle(slide, "So sánh với kết quả công bố", 29);
  deleteContentPlaceholders(slide);
  addTable(
    slide,
    [
      ["Metric", "Bài báo", "Chạy thử local"],
      ["AUC", "0.99913", "0.999831"],
      ["High confidence", "6622/7179 = 92.2%", "996/1077 = 92.48%"],
      ["False Positive", "0", "0"],
      ["False Negative", "44/7179 = 0.61%", "4/1077 = 0.37%"],
      ["Accuracy", "99.4%", "99.63%"],
    ],
    62,
    126,
    [230, 292, 292],
    52,
    { fontSize: 16, boldFirstCol: true },
  );
  addCard(
    slide,
    "Nhận xét",
    "Kết quả chạy thử bám rất sát mục tiêu của bài báo: tỷ lệ High confidence khoảng 92%, FP = 0. AUC local cao hơn nhẹ do split ngẫu nhiên và pipeline sklearn có thể khác đúng cấu hình gốc.",
    120,
    486,
    720,
    94,
    { fill: LIGHT_BLUE, titleColor: BLUE, bodySize: 16.5 },
  );
  addSmallSource(slide, "Bài báo: BMC Genomics 2018, 19:263; PMID 29665779; PMCID PMC5904977");
}

// Slide 8: bioinformatics interpretation
{
  const slide = slideAt(7);
  addContentChrome(slide);
  setTitle(slide, "Ý nghĩa Bioinformatics của các feature", 28);
  deleteContentPlaceholders(slide);
  addCard(slide, "Read support", "DP/AD/AF phản ánh số read ủng hộ variant và độ cân bằng allele.", 58, 130, 380, 112, {
    fill: LIGHT_BLUE,
    titleColor: BLUE,
    bodySize: 17,
  });
  addCard(slide, "Call quality", "QUAL/QD/GQ/MQ giúp phân biệt tín hiệu thật với lỗi gọi hoặc lỗi alignment.", 58, 272, 380, 112, {
    fill: LIGHT_GREEN,
    titleColor: GREEN,
    bodySize: 17,
  });
  addCard(slide, "Artifact context", "FS/WHR/HPL/GC bắt các vùng khó: strand bias, homopolymer, GC extreme.", 58, 414, 380, 112, {
    fill: LIGHT_RED,
    titleColor: RED,
    bodySize: 17,
  });
  addCard(
    slide,
    "Vì sao Logistic Regression hợp lý",
    "- Dữ liệu ít chiều, feature có ý nghĩa sinh học rõ.\n- StandardScaler giúp hệ số so sánh được.\n- Ngưỡng bảo thủ biến xác suất thành chính sách báo cáo.\n- Có thể kiểm tra coefficient để giải thích cho thầy/hội đồng.",
    506,
    160,
    360,
    282,
    { fill: "#FFFFFF", bodySize: 17 },
  );
}

// Slide 9: limitations and next steps
{
  const slide = slideAt(8);
  addContentChrome(slide);
  setTitle(slide, "Thiếu sót hiện tại và bước tiếp theo", 29);
  deleteContentPlaceholders(slide);
  addCard(
    slide,
    "Thiếu sót cần nói rõ",
    "- Split local có thể khác bài báo gốc.\n- Chưa có external validation từ assay/phòng lab khác.\n- Threshold zero-FP trên Dev vẫn có rủi ro overfit.\n- Chưa phân tích coefficient và calibration curve.\n- Class imbalance: negative ít hơn positive.",
    62,
    126,
    390,
    320,
    { fill: "#FFFFFF", titleColor: RED, bodySize: 17 },
  );
  addCard(
    slide,
    "Bước tiếp theo đề xuất",
    "- Cố định preprocessing pipeline và random seed.\n- Báo cáo coefficient/odds ratio theo feature.\n- Nested CV hoặc bootstrap confidence interval.\n- Test trên dữ liệu variant call nội bộ.\n- So sánh baseline hard filter vs ML threshold.",
    508,
    126,
    390,
    320,
    { fill: "#FFFFFF", titleColor: BLUE, bodySize: 17 },
  );
  addBox(slide, 110, 505, 740, 64, { fill: LIGHT_GOLD, stroke: "#E4C989" });
  addText(slide, "Thông điệp khi báo cáo: model này là công cụ giảm xác nhận bổ sung, không thay thế quy trình kiểm định lâm sàng.", 136, 516, 688, 42, {
    fontSize: 16.5,
    color: AMBER,
    bold: true,
    align: "center",
  });
}

// Slide 10: closing
{
  const slide = slideAt(9);
  clearTemplatePlaceholders(slide);
  addClosingChrome(slide);
  const title = shapeByName(slide, "Title");
  setExistingText(title, "CẢM ƠN THẦY\nQ&A", {
    fontSize: 34,
    color: RED,
    bold: true,
    typeface: "Lato",
    align: "center",
    valign: "middle",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
  });
  if (title?.bringToFront) title.bringToFront();
  addText(slide, "File chạy thử: D:\\A\\train_conservative_threshold.py\nDữ liệu: D:\\A\\12864_2018_4659_MOESM2_ESM.xlsx", 410, 438, 510, 50, {
    fontSize: 14.5,
    color: INK,
    align: "center",
  });
}

await fs.mkdir(previewDir, { recursive: true });
await fs.mkdir(layoutDir, { recursive: true });
const previewPaths = [];
const layoutPaths = [];

for (let i = 0; i < presentation.slides.items.length; i += 1) {
  const slide = presentation.slides.getItem(i);
  const preview = await presentation.export({ slide, format: "png", scale: 1 });
  const previewPath = path.join(previewDir, `slide-${String(i + 1).padStart(2, "0")}.png`);
  await saveBlobToFile(preview, previewPath);
  previewPaths.push(previewPath);

  const layoutBlob = await presentation.export({ slide, format: "layout" });
  const layoutPath = path.join(layoutDir, `slide-${String(i + 1).padStart(2, "0")}.layout.json`);
  await fs.writeFile(layoutPath, await layoutBlob.text(), "utf8");
  layoutPaths.push(layoutPath);
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(outputPptx);
const stat = await fs.stat(outputPptx);

const manifest = {
  output: outputPptx,
  outputBytes: stat.size,
  previewDir,
  previewPaths,
  layoutDir,
  layoutPaths,
  slideCount: presentation.slides.items.length,
};
await fs.writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
console.log(JSON.stringify(manifest, null, 2));
