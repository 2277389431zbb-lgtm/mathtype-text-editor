using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MathTypeEditor
{
    internal static class EditorTests
    {
        internal static int Run(bool integration)
        {
            var log = new List<string>(); int checks = 0;
            Action<bool, string> check = delegate(bool ok, string name) { if (!ok) throw new Exception("FAIL: " + name); checks++; log.Add("PASS: " + name); };
            string logPath = Path.Combine(EditorForm.BaseDir, integration ? "v2_集成验证.txt" : "v2_自检结果.txt");
            try
            {
                List<string> warnings;
                string text = "正文加粗不能变。行内 \\(E=2.06\\times10^8\\) 文字接着。\n" + EditorForm.DemoFormula + "\n结尾 $x_i^2$。";
                var scan = FormulaText.Scan(text, out warnings);
                check(scan.Count == 3, "mixed prose recognizes three formulas");
                check(warnings.Count == 0, "mixed text has no parse warnings");
                check(scan[0].Display == false && scan[1].Display && scan[2].Display == false, "inline and display distinguished");
                check(scan[1].Number == "5-24", "tag extracted without changing body");
                check(text.Substring(scan[0].Start, scan[0].Length) == @"\(E=2.06\times10^8\)", "replacement bounds exclude adjacent text");
                check(FormulaText.Scan("正文\n[\nx^2=1\n]\n尾文", out warnings).Count == 1, "plain square bracket blocks");
                check(FormulaText.Scan("正文\n[\vx^2=1\v]\n尾文", out warnings).Count == 1, "Word soft line breaks recognized");
                check(FormulaText.Scan(@"钢材取 (E=2.06\times10^8\ \mathrm{kN/m^2})，则：", out warnings).Count == 1, "plain parentheses containing TeX recognized");
                check(FormulaText.Parse(@"(x+1)\times(y+1)", false).Body == @"(x+1)\times(y+1)", "mathematical grouping parentheses preserved");
                check(FormulaText.Scan("价格$100，另一个$200", out warnings).Count == 0, "ordinary currency not mistaken for a formula");
                check(FormulaText.Scan(@"价格 \$100 和 \$200。", out warnings).Count == 0, "escaped currency preserved");
                check(FormulaText.Scan(@"未完成 \[x^2", out warnings).Count == 0, "incomplete formula untouched");
                check(FormulaText.Scan(@"前文 \(x_{1\) 后文", out warnings).Count == 0 && warnings.Count == 1, "malformed braces retained with warning");
                check(FormulaText.Scan("```latex\n\\frac{1}{2}\n```", out warnings).Count == 1, "latex fenced blocks");
                check(FormulaText.Scan(@"\begin{equation}a=b\tag{2-1}\end{equation}", out warnings)[0].Number == "2-1", "equation environments");
                string katex = "<p><strong style='color:#cc0000'>原文</strong><span class='katex'><span class='katex-mathml'><math><semantics><mi>x</mi><annotation encoding='application/x-tex'>x^2</annotation></semantics></math></span><span class='katex-html'>duplicate</span></span>尾文</p>";
                string html = HtmlMath.Prepare(katex);
                check(html.Contains(@"\(x^2\)") && !html.Contains("duplicate"), "KaTeX source extracted exactly once");
                check(html.Contains("<strong style='color:#cc0000'>原文</strong>") && html.Contains("尾文"), "non-math HTML styles and text retained");
                check(HtmlMath.Prepare("<p><math><mfrac><mn>1</mn><mn>2</mn></mfrac></math></p>").Contains(@"\frac{1}{2}"), "MathML fractions supported");
                check(HtmlMath.Prepare("<math><semantics><mi>x</mi><annotation encoding='application/x-tex'>x\\tag{3}</annotation></semantics></math>").Contains(@"\tag{3}"), "annotation preferred over visual MathML");
                check(!HtmlMath.Prepare("<p onclick='bad()'>正文</p><script>bad()</script>").Contains("bad()"), "active HTML removed");
                if (integration)
                {
                    using (var box = new RichTextBox())
                    using (var engine = new WordEngine())
                    {
                        var handle = box.Handle; EditorForm.BuildDemo(box); string before = box.Text; int lines = box.Lines.Length;
                        int bodyStart = box.Text.IndexOf("冠梁", StringComparison.Ordinal); box.Select(bodyStart, 0); box.SelectionIndent = 18; box.SelectionAlignment = HorizontalAlignment.Center;
                        var formulas = FormulaText.Scan(box.Text, out warnings);
                        check(formulas.Count == 2 && warnings.Count == 0, "formatted demo contains two valid formulas");
                        var results = new List<Rendered>();
                        foreach (var f in formulas) results.Add(new Rendered { Formula = f, Result = engine.Render(f) });
                        check(results.All(r => r.Result.ProgId.Contains("DSMT")), "backend produces real MathType OLE objects");
                        EditorForm.Apply(box, results);
                        check(!box.Text.Contains(@"\frac") && box.Text.Contains("(5-24)"), "render replaces only equation code and retains number");
                        check(box.Lines.Length == lines, "inline formulas do not add paragraphs");
                        box.Select(box.Text.IndexOf("冠梁", StringComparison.Ordinal), 0); check(box.SelectionIndent == 18 && box.SelectionAlignment == HorizontalAlignment.Center, "paragraph indentation and alignment preserved");
                        box.Select(0, 7); check(box.SelectionFont != null && box.SelectionFont.Bold && Math.Abs(box.SelectionFont.Size - 16) < .1, "heading font and bold survive replacements");
                        check(box.SelectionColor == Color.FromArgb(20, 85, 140), "heading color preserved");
                        int conclusion = box.Text.IndexOf("结论：", StringComparison.Ordinal); box.Select(conclusion, 3);
                        check(box.SelectionFont.Bold && box.SelectionColor == Color.FromArgb(180, 45, 35), "formatting after equations preserved");
                        string all = engine.VerifyRtf(box.Rtf, 2);
                        check(all.Contains("冠梁与支撑采用下列计算关系。行内公式") && all.Contains("与文字一起保留。") && all.Contains("文字可以继续编辑。"), "Word receives surrounding Chinese prose unchanged");
                        TextStyle style = engine.InspectStyle(box.Rtf, "钢支撑刚度计算");
                        check(style.Bold == -1 && Math.Abs(style.Size - 16) < .1 && style.Color == (20 + 85 * 256 + 140 * 65536), "Word receives original heading formatting");
                        using (var reopened = new RichTextBox()) { var h = reopened.Handle; reopened.Rtf = box.Rtf; engine.VerifyRtf(reopened.Rtf, 2); check(true, "draft reload retains both MathType objects"); }
                        string imported = engine.HtmlToRtf(katex);
                        using (var rich = new RichTextBox())
                        { var h = rich.Handle; rich.Rtf = imported; var math = FormulaText.Scan(rich.Text, out warnings); check(math.Count == 1, "browser rich paste yields one formula"); rich.Select(0,2); check(rich.SelectionFont.Bold && rich.SelectionColor.R == 204, "browser bold and color preserved"); }
                        string testDir = Path.Combine(Path.GetTempPath(), "MathTypeEditor-v2-tests"); Directory.CreateDirectory(testDir); box.SaveFile(Path.Combine(testDir, "验证结果.rtf"));
                        log.Add("RTF preview: " + Path.Combine(testDir, "验证结果.rtf"));
                    }
                }
                log.Add("ALL PASS: " + checks + " checks"); File.WriteAllLines(logPath, log, Encoding.UTF8); return 0;
            }
            catch (Exception e) { log.Add(e.ToString()); File.WriteAllLines(logPath, log, Encoding.UTF8); return 1; }
        }
    }
}
