using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MathTypeEditor
{
    internal sealed class RichEditor : RichTextBox
    {
        internal event EventHandler PasteRequested;
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0302 && PasteRequested != null) { PasteRequested(this, EventArgs.Empty); return; }
            base.WndProc(ref m);
        }
        protected override bool ProcessCmdKey(ref Message m, Keys key)
        {
            if (key == (Keys.Control | Keys.V) || key == (Keys.Shift | Keys.Insert)) { if (PasteRequested != null) PasteRequested(this, EventArgs.Empty); return true; }
            return base.ProcessCmdKey(ref m, key);
        }
    }

    internal sealed class Rendered
    {
        internal FormulaSpec Formula;
        internal FormulaRender Result;
    }
    internal sealed class UndoBatch { internal string Before, After; }

    internal sealed class EditorForm : Form
    {
        internal readonly RichEditor Editor;
        readonly WordWorker worker = new WordWorker();
        readonly Label status;
        readonly ComboBox targets, fonts, sizes;
        readonly CheckBox automatic;
        readonly FlowLayoutPanel actions, format;
        readonly Button insert, cancel;
        readonly System.Windows.Forms.Timer autosave;
        readonly Stack<UndoBatch> conversions = new Stack<UndoBatch>();
        readonly string draftDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MathTypeOneClick");
        string notes = "";
        bool busy, dirty, restoring;
        CancellationTokenSource cancellation;
        internal static readonly string BaseDir = AppDomain.CurrentDomain.BaseDirectory;
        internal const string DemoFormula = @"\[k_{7.5}=\frac{2.06\times10^8\times0.0298074}{7.5}\approx8.19\times10^5\ \mathrm{kN/m}\tag{5-24}\]";

        internal EditorForm()
        {
            Text = "MathType 文本编辑器 v2";
            Font = new Font("Microsoft YaHei UI", 10F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(920, 580); MinimumSize = new Size(760, 520);
            StartPosition = FormStartPosition.CenterScreen; Icon = SystemIcons.Application;
            BackColor = Color.FromArgb(245, 247, 249);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16, 10, 16, 10), ColumnCount = 1, RowCount = 6 };
            foreach (int height in new[] { 58, 44, 42 }) root.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            var header = new Label { Text = "正文与公式一起编辑\r\n粘贴整段内容，自动转换公式；编辑完成后插回 Word。", Dock = DockStyle.Fill, Font = new Font(Font.FontFamily, 11F), ForeColor = Color.FromArgb(26, 75, 85), Padding = new Padding(0, 3, 0, 0) };
            root.Controls.Add(header, 0, 0);
            actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            actions.Controls.Add(MakeButton("粘贴内容", async delegate { await PasteContent(); }, 112));
            actions.Controls.Add(MakeButton("转换剩余公式", async delegate { await ConvertContent(); }, 135));
            actions.Controls.Add(MakeButton("撤销本次转换", delegate { RestoreConversion(); }, 135));
            actions.Controls.Add(MakeButton("保存草稿", delegate { SaveDraftAs(); }, 110));
            actions.Controls.Add(MakeButton("打开草稿", async delegate { await OpenDraft(); }, 110));
            actions.Controls.Add(MakeButton("示例", async delegate { await ShowDemo(); }, 75));
            actions.Controls.Add(MakeButton("说明", delegate { Process.Start(Path.Combine(BaseDir, "使用说明.txt")); }, 65));
            root.Controls.Add(actions, 0, 1);
            format = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            fonts = new ComboBox { Width = 135, DropDownStyle = ComboBoxStyle.DropDown, Text = "宋体", Margin = new Padding(3, 5, 7, 0) };
            fonts.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Microsoft YaHei", "Times New Roman", "Arial", "Calibri" });
            fonts.SelectionChangeCommitted += delegate { ChangeFont(); }; fonts.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { ChangeFont(); e.SuppressKeyPress = true; } };
            sizes = new ComboBox { Width = 60, Text = "12", Margin = new Padding(3, 5, 8, 0) };
            sizes.Items.AddRange(new object[] { "9", "10", "10.5", "11", "12", "14", "16", "18", "22", "24" });
            sizes.SelectionChangeCommitted += delegate { ChangeFont(); }; sizes.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { ChangeFont(); e.SuppressKeyPress = true; } };
            format.Controls.Add(fonts); format.Controls.Add(sizes);
            format.Controls.Add(MakeButton("粗体", delegate { ToggleStyle(FontStyle.Bold); }, 62));
            format.Controls.Add(MakeButton("斜体", delegate { ToggleStyle(FontStyle.Italic); }, 62));
            format.Controls.Add(MakeButton("下划线", delegate { ToggleStyle(FontStyle.Underline); }, 75));
            format.Controls.Add(MakeButton("颜色", delegate { using (var dialog = new ColorDialog { Color = Editor.SelectionColor }) if (dialog.ShowDialog(this) == DialogResult.OK) Editor.SelectionColor = dialog.Color; }, 62));
            format.Controls.Add(MakeButton("左对齐", delegate { Editor.SelectionAlignment = HorizontalAlignment.Left; Editor.Focus(); }, 75));
            format.Controls.Add(MakeButton("居中", delegate { Editor.SelectionAlignment = HorizontalAlignment.Center; Editor.Focus(); }, 62));
            automatic = new CheckBox { Text = "粘贴后自动转换", Checked = true, AutoSize = true, Margin = new Padding(8, 10, 0, 0) }; format.Controls.Add(automatic);
            root.Controls.Add(format, 0, 2);
            Editor = new RichEditor { Dock = DockStyle.Fill, Font = new Font("宋体", 12F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, DetectUrls = false, HideSelection = false, EnableAutoDragDrop = true, AcceptsTab = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Both, AccessibleName = "正文与 MathType 公式编辑区" };
            Editor.PasteRequested += async delegate { await PasteContent(); };
            Editor.TextChanged += delegate { if (!restoring) dirty = true; };
            Editor.SelectionChanged += delegate { if (!busy) { Font f = Editor.SelectionFont; if (f != null) { fonts.Text = f.Name; sizes.Text = f.Size.ToString("0.##"); } } };
            var context = new ContextMenuStrip();
            context.Items.Add("剪切", null, delegate { Editor.Cut(); }); context.Items.Add("复制", null, delegate { Editor.Copy(); });
            context.Items.Add("粘贴并保留格式", null, async delegate { await PasteContent(); }); context.Items.Add("全选", null, delegate { Editor.SelectAll(); });
            Editor.ContextMenuStrip = context;
            root.Controls.Add(Editor, 0, 3);
            var targetRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Padding = new Padding(0, 6, 0, 0) };
            targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92)); targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85)); targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 165));
            targetRow.Controls.Add(new Label { Text = "插入目标", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            targets = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 8, 0) }; targetRow.Controls.Add(targets, 1, 0);
            targetRow.Controls.Add(MakeButton("刷新", delegate { RefreshTargets(); }, 75), 2, 0);
            insert = MakeButton("插入回 Word", async delegate { await InsertBack(); }, 155); insert.BackColor = Color.FromArgb(19, 104, 113); insert.ForeColor = Color.White; insert.FlatStyle = FlatStyle.Flat; targetRow.Controls.Add(insert, 3, 0);
            root.Controls.Add(targetRow, 0, 4);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88)); footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
            status = new Label { Dock = DockStyle.Fill, Text = "在编辑区按 Ctrl+V，或点击“粘贴内容”。保留剪贴板中已有的文字格式。", ForeColor = Color.FromArgb(65, 79, 88), Padding = new Padding(0, 7, 0, 0) }; footer.Controls.Add(status, 0, 0);
            cancel = MakeButton("停止转换", delegate { if (cancellation != null) cancellation.Cancel(); }, 85); cancel.Enabled = false; footer.Controls.Add(cancel, 1, 0);
            footer.Controls.Add(MakeButton("查看提示", delegate { MessageBox.Show(this, notes.Length > 0 ? notes : "没有转换错误。支持带边界的 LaTeX，以及含公式源码的浏览器富文本。", "转换提示"); }, 85), 2, 0);
            root.Controls.Add(footer, 0, 5); Controls.Add(root);
            autosave = new System.Windows.Forms.Timer { Interval = 7000 }; autosave.Tick += delegate { if (dirty && !busy) SaveRecovery(); }; autosave.Start();
            Shown += delegate
            {
                var area = Screen.FromControl(this).WorkingArea; if (Height > area.Height - 20) Height = area.Height - 20; if (Width > area.Width - 20) Width = area.Width - 20;
                RefreshTargets(); RestoreRecovery(); Editor.Focus();
            };
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (busy) { e.Cancel = true; if (cancellation != null) cancellation.Cancel(); status.Text = "正在结束当前操作，完成后可关闭窗口。"; return; } SaveRecovery(); autosave.Stop(); worker.Dispose(); };
        }
        Button MakeButton(string text, EventHandler action, int width)
        { var b = new Button { Text = text, Width = width, Height = 32, Margin = new Padding(3, 3, 5, 3), UseVisualStyleBackColor = true }; b.Click += action; return b; }
        void SetBusy(bool value)
        { busy = value; Editor.ReadOnly = value; actions.Enabled = !value; format.Enabled = !value; insert.Enabled = !value; targets.Enabled = !value; cancel.Enabled = value; Cursor = value ? Cursors.AppStarting : Cursors.Default; }
        void Error(Exception e)
        {
            notes = e.Message; status.Text = "未完成：" + e.Message;
            try { File.AppendAllText(Path.Combine(BaseDir, "编辑器运行记录.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + e.GetType().Name + " " + e.Message + Environment.NewLine, Encoding.UTF8); } catch { }
        }
        void RefreshTargets()
        {
            var old = targets.SelectedItem as WordTarget; var list = Office.Targets(); targets.Items.Clear(); foreach (var item in list) targets.Items.Add(item);
            if (old != null) for (int i = 0; i < targets.Items.Count; i++) if (((WordTarget)targets.Items[i]).Handle == old.Handle) targets.SelectedIndex = i;
            if (targets.SelectedIndex < 0 && targets.Items.Count > 0) targets.SelectedIndex = 0;
        }
        async Task PasteContent()
        {
            if (busy) return;
            try
            {
                // RTF is preferred over HTML to keep existing MathType objects and Word formatting.
                if (Clipboard.ContainsData(DataFormats.Rtf)) Editor.SelectedRtf = (string)Clipboard.GetData(DataFormats.Rtf);
                else if (Clipboard.ContainsData(DataFormats.Html))
                {
                    string html = Convert.ToString(Clipboard.GetData(DataFormats.Html));
                    SetBusy(true); status.Text = "正在保留浏览器中的文字格式，并读取公式源码……";
                    try { Editor.SelectedRtf = await worker.Run(e => e.HtmlToRtf(html)); } finally { SetBusy(false); }
                }
                else if (Clipboard.ContainsText()) Editor.SelectedText = Clipboard.GetText(TextDataFormat.UnicodeText);
                else { status.Text = "剪贴板没有可编辑的文字。公式截图不能直接转换为 MathType。"; return; }
                Editor.Focus(); dirty = true;
                if (automatic.Checked) await ConvertContent(); else status.Text = "已保留格式粘贴。可继续编辑，或点击“转换剩余公式”。";
                SaveRecovery();
            }
            catch (Exception e) { SetBusy(false); Error(e); }
        }
        internal static string Fragment(string rtf)
        {
            using (var scratch = new RichTextBox())
            {
                var handle = scratch.Handle; scratch.Rtf = rtf;
                int length = scratch.TextLength; while (length > 0 && (scratch.Text[length - 1] == '\n' || scratch.Text[length - 1] == '\r')) length--;
                scratch.Select(0, length); return scratch.SelectedRtf;
            }
        }
        internal static void Apply(RichTextBox box, IList<Rendered> results)
        {
            foreach (var item in results.OrderByDescending(r => r.Formula.Start))
            { box.Select(item.Formula.Start, item.Formula.Length); box.SelectedRtf = Fragment(item.Result.Rtf); }
        }
        async Task ConvertContent()
        {
            if (busy) return;
            List<string> warnings; var formulas = FormulaText.Scan(Editor.Text, out warnings);
            if (formulas.Count == 0) { notes = String.Join(Environment.NewLine, warnings); status.Text = warnings.Count > 0 ? "发现不完整的公式，已保留代码。点击“查看提示”。" : "没有待转换公式。已转换的 MathType 对象会保持原样。"; return; }
            string original = Editor.Rtf; var result = new List<Rendered>();
            cancellation = new CancellationTokenSource(); SetBusy(true);
            try
            {
                for (int i = 0; i < formulas.Count; i++)
                {
                    if (cancellation.IsCancellationRequested) break;
                    FormulaSpec f = formulas[i]; status.Text = "正在转换公式 " + (i + 1) + "/" + formulas.Count + "；正文格式保持在编辑区中。";
                    try { FormulaRender r = await worker.Run(e => e.Render(f)); result.Add(new Rendered { Formula = f, Result = r }); }
                    catch (Exception e) { warnings.Add("公式 " + (i + 1) + "：" + e.Message); }
                }
                if (result.Count > 0)
                {
                    Apply(Editor, result); if (conversions.Count >= 8) conversions.Clear(); conversions.Push(new UndoBatch { Before = original, After = Editor.Rtf }); Editor.Select(Math.Min(Editor.TextLength, result[0].Formula.Start), 0); dirty = true;
                }
                notes = String.Join(Environment.NewLine, warnings);
                status.Text = "已转换 " + result.Count + " 条公式" + (warnings.Count > 0 ? "，未完成的保留代码；详见提示。" : "。可继续编辑，或点击“插入回 Word”。");
                if (cancellation.IsCancellationRequested) status.Text = "已停止；完成 " + result.Count + " 条，其余公式保留代码。";
                SaveRecovery();
            }
            catch (Exception e) { Editor.Rtf = original; Error(e); }
            finally { cancellation.Dispose(); cancellation = null; SetBusy(false); Editor.Focus(); }
        }
        async Task InsertBack()
        {
            if (busy) return;
            var target = targets.SelectedItem as WordTarget;
            if (target == null) { RefreshTargets(); status.Text = "请先在 Microsoft Word 中打开目标文档，并在下方选择它。"; return; }
            if (Editor.TextLength == 0 && !Editor.Rtf.Contains(@"\object")) { status.Text = "编辑区为空，请先粘贴内容。"; return; }
            // Finish any newly typed TeX before sending the edited result back.
            await ConvertContent(); if (busy) return;
            string rtf = Editor.Rtf; SetBusy(true); status.Text = "正在插入目标 Word 文档……";
            try { status.Text = await worker.Run(e => e.Insert(rtf, target.Handle)); if (notes.Length > 0) status.Text += " 部分公式保留代码，详见提示。"; SaveRecovery(); }
            catch (Exception e) { Error(e); }
            finally { SetBusy(false); }
        }
        void RestoreConversion()
        { if (busy) return; if (conversions.Count == 0) { status.Text = "没有可撤销的公式转换。文字编辑可在编辑区按 Ctrl+Z。"; return; } if (Editor.Rtf != conversions.Peek().After) { status.Text = "转换后已有新的编辑。为保留这些修改，请在编辑区用 Ctrl+Z 逐步撤销。"; return; } Editor.Rtf = conversions.Pop().Before; dirty = true; status.Text = "已撤销最近一批公式转换。"; Editor.Focus(); }
        void ChangeFont()
        { float size; if (!Single.TryParse(sizes.Text, out size) || size < 5 || size > 100) return; try { Font current = Editor.SelectionFont ?? Editor.Font; Editor.SelectionFont = new Font(fonts.Text, size, current.Style); Editor.Focus(); } catch (Exception e) { Error(e); } }
        void ToggleStyle(FontStyle style)
        { Font f = Editor.SelectionFont ?? Editor.Font; Editor.SelectionFont = new Font(f, f.Style ^ style); Editor.Focus(); }
        void SaveRecovery()
        { try { Directory.CreateDirectory(draftDir); File.WriteAllText(Path.Combine(draftDir, "draft.rtf"), Editor.Rtf, Encoding.ASCII); dirty = false; } catch { } }
        void RestoreRecovery()
        { string path = Path.Combine(draftDir, "draft.rtf"); if (!File.Exists(path)) return; try { restoring = true; Editor.LoadFile(path); restoring = false; dirty = false; status.Text = "已恢复上次本地草稿，可继续编辑。"; } catch { restoring = false; } }
        void SaveDraftAs()
        { using (var dialog = new SaveFileDialog { Filter = "富文本草稿（含 MathType 对象）|*.rtf", FileName = "公式编辑草稿.rtf" }) if (dialog.ShowDialog(this) == DialogResult.OK) { Editor.SaveFile(dialog.FileName, RichTextBoxStreamType.RichText); status.Text = "草稿已保存，可用本软件或 Word 继续打开。"; } }
        async Task OpenDraft()
        { using (var dialog = new OpenFileDialog { Filter = "富文本草稿|*.rtf" }) if (dialog.ShowDialog(this) == DialogResult.OK) { try { SaveRecovery(); Editor.LoadFile(dialog.FileName); conversions.Clear(); dirty = true; if (automatic.Checked) await ConvertContent(); } catch (Exception e) { Error(e); } } }
        async Task ShowDemo()
        {
            if (Editor.TextLength > 0) { using (var demo = new RichTextBox()) { BuildDemo(demo); Editor.Select(Editor.TextLength, 0); Editor.SelectedText = "\n"; Editor.SelectedRtf = demo.Rtf; } }
            else BuildDemo(Editor);
            dirty = true; await ConvertContent();
        }
        internal static void BuildDemo(RichTextBox box)
        {
            box.Clear(); box.Font = new Font("宋体", 12F);
            box.AppendText("钢支撑刚度计算\n"); box.Select(0, 8); box.SelectionFont = new Font("黑体", 16F, FontStyle.Bold); box.SelectionColor = Color.FromArgb(20, 85, 140);
            box.Select(box.TextLength, 0); box.SelectionFont = new Font("宋体", 12F); box.SelectionColor = Color.Black;
            box.AppendText("冠梁与支撑采用下列计算关系。行内公式 \\(E=2.06\\times10^8\\) 与文字一起保留。\n");
            box.AppendText(DemoFormula + "\n");
            box.AppendText("结论：计算结果保持不变，文字可以继续编辑。\n");
            int start = box.Text.IndexOf("结论：", StringComparison.Ordinal); box.Select(start, 3); box.SelectionFont = new Font("宋体", 12F, FontStyle.Bold); box.SelectionColor = Color.FromArgb(180, 45, 35); box.Select(box.TextLength, 0);
        }
    }

    internal static class Program
    {
        [STAThread] static int Main(string[] args)
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length > 0 && args[0] == "--self-test") return EditorTests.Run(false);
            if (args.Length > 0 && args[0] == "--integration-test") return EditorTests.Run(true);
            bool created;
            using (var mutex = new Mutex(true, @"Local\MathTypeTextEditorV2", out created))
            {
                if (!created) { MessageBox.Show("编辑器已经打开，请切换到 MathType 文本编辑器窗口。", "MathType 文本编辑器"); return 0; }
                Application.Run(new EditorForm());
            }
            return 0;
        }
    }
}
