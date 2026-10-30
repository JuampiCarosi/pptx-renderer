import PptxGenJS from "pptxgenjs-plus"; // 4.3.4, pptxgenjs-compatible API

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
const common = {
  barDir: "bar",
  barGrouping: "clustered",
  chartColors: ["1F4E79"],
  // Hidden value axis with explicit max: headroom for the labels at the bar ends.
  valAxisHidden: true,
  valAxisLineShow: false,
  valGridLine: { style: "none" },
  catGridLine: { style: "none" },
  // Category axis line explicitly hidden (a:ln/a:noFill).
  catAxisLineShow: false,
  catAxisMajorTickMark: "none",
  catAxisLabelFontSize: 14,
  catAxisLabelPos: "low",
  showLegend: false,
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelFontSize: 14,
  dataLabelFormatCode: "#,##0",
  // White plot area without an outline.
  plotArea: { fill: { color: "FFFFFF" }, border: { pt: 0, color: "FFFFFF" } },
};
const slide = pptx.addSlide();
slide.addText("Horizontal bars · dLblPos=outEnd · hidden valAx with max · catAx ln noFill", { x: 0.5, y: 0.3, w: 12.3, h: 0.5, fontSize: 18, bold: true });
const values = [775, 770, 581, 560, 551, 499, 471, 322];
slide.addChart(pptx.ChartType.bar, [{ name: "Value", labels: ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"], values }], {
  ...common, x: 0.5, y: 1.0, w: 6.0, h: 5.8, valAxisMinVal: 0, valAxisMaxVal: Math.max(...values) * 1.24,
});
const signed = [420, -310, 180, -95];
slide.addChart(pptx.ChartType.bar, [{ name: "Change", labels: ["North", "South", "East", "West"], values: signed }], {
  ...common, x: 6.9, y: 1.0, w: 6.0, h: 5.8, valAxisMinVal: -310 * 1.4, valAxisMaxVal: 420 * 1.4,
});
await pptx.writeFile({ fileName: "horizontal-bar-labels.pptx" });
console.log("ok");
