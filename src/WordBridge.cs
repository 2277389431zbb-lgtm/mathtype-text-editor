using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MathTypeEditor
{
    internal static class Office
    {
        internal delegate bool WindowCallback(IntPtr h, IntPtr p);
        [DllImport("user32.dll")] internal static extern bool EnumWindows(WindowCallback cb, IntPtr p);
        [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr h, WindowCallback cb, IntPtr p);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr h, StringBuilder text, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetWindowText(IntPtr h, StringBuilder text, int size);
        [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object obj);
        internal static volatile int PrivatePid;

        internal static object WindowObject(IntPtr root)
        {
            object result = null;
            EnumChildWindows(root, delegate(IntPtr h, IntPtr p)
            {
                var c = new StringBuilder(128); GetClassName(h, c, c.Capacity);
                if (c.ToString() == "_WwG")
                {
                    Guid iid = new Guid("00020400-0000-0000-C000-000000000046"); object obj = null;
                    if (AccessibleObjectFromWindow(h, 0xFFFFFFF0, ref iid, ref obj) >= 0 && obj != null) { result = obj; return false; }
                }
                return true;
            }, IntPtr.Zero);
            if (result == null) throw new Exception("Word 编辑区尚未就绪。请先打开目标文档，并关闭 Word 中的提示框。");
            return result;
        }
        internal static object ProcessObject(int pid)
        {
            object result = null;
            EnumWindows(delegate(IntPtr h, IntPtr p) { uint n; GetWindowThreadProcessId(h, out n); if (n == (uint)pid) try { result = WindowObject(h); return false; } catch { } return true; }, IntPtr.Zero);
            if (result == null) throw new Exception("后台 Word 尚未就绪。");
            return result;
        }
        internal static List<WordTarget> Targets()
        {
            var result = new List<WordTarget>();
            EnumWindows(delegate(IntPtr h, IntPtr p)
            {
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (!IsWindowVisible(h) || pid == PrivatePid) return true;
                try
                {
                    if (!Process.GetProcessById((int)pid).ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return true;
                    var text = new StringBuilder(512); GetWindowText(h, text, text.Capacity);
                    var cls = new StringBuilder(128); GetClassName(h, cls, cls.Capacity);
                    if (cls.ToString() == "OpusApp" && text.Length > 0) result.Add(new WordTarget { Handle = h, Title = text.ToString() });
                }
                catch { }
                return true;
            }, IntPtr.Zero);
            return result;
        }
        internal static string WordExe()
        {
            foreach (var process in Process.GetProcessesByName("WINWORD"))
                try { if (File.Exists(process.MainModule.FileName)) return process.MainModule.FileName; } catch { }
            foreach (string root in new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) })
                foreach (string sub in new[] { @"Microsoft Office\root\Office16\WINWORD.EXE", @"Microsoft Office\Office16\WINWORD.EXE", @"Microsoft Office\Office15\WINWORD.EXE" })
                { string path = Path.Combine(root, sub); if (File.Exists(path)) return path; }
            throw new Exception("未找到 Microsoft Word。请先安装或打开 Microsoft Word。");
        }
    }
    internal sealed class WordTarget
    {
        internal IntPtr Handle; internal string Title;
        public override string ToString() { return Title; }
    }

    internal sealed class FormulaRender
    {
        internal string Rtf;
        internal string ProgId;
    }
    internal sealed class TextStyle
    {
        internal string Font; internal float Size; internal int Bold, Color;
    }

    internal sealed class WordEngine : IDisposable
    {
        dynamic word;
        Process process;
        readonly string session = Path.Combine(Path.GetTempPath(), "MathTypeEditor", Guid.NewGuid().ToString("N"));
        internal void Ensure()
        {
            if (word != null) { try { var v = word.Version; return; } catch { word = null; } }
            Directory.CreateDirectory(session);
            string blank = NewPath(".rtf"); File.WriteAllText(blank, @"{\rtf1\ansi }", Encoding.ASCII);
            process = Process.Start(new ProcessStartInfo(Office.WordExe(), "/x /q /automation \"" + blank + "\"") { WindowStyle = ProcessWindowStyle.Hidden, UseShellExecute = false });
            Office.PrivatePid = process.Id;
            for (int i = 0; i < 120; i++)
            {
                Thread.Sleep(100);
                if (process.HasExited) break;
                try { dynamic win = Office.ProcessObject(process.Id); word = win.Application; word.Visible = false; word.DisplayAlerts = 0; return; } catch { }
            }
            throw new Exception("后台 Word 启动超时。请先完成 Word 中的首次启动或许可提示，再重试。");
        }
        string NewPath(string suffix) { Directory.CreateDirectory(session); return Path.Combine(session, Guid.NewGuid().ToString("N") + suffix); }
        static void CloseDocument(dynamic doc) { if (doc != null) try { doc.Close(0); } catch { } }
        static void Remove(string path) { try { if (File.Exists(path)) File.Delete(path); } catch { } }
        internal FormulaRender Render(FormulaSpec formula)
        {
            Ensure(); dynamic doc = null;
            string output = NewPath(".rtf");
            try
            {
                doc = word.Documents.Add();
                string tex = "$" + (formula.Display ? @"\displaystyle " : "") + formula.Body + "$";
                doc.Content.Text = tex;
                doc.Range(0, tex.Length).Select();
                try { word.Run("MTCommand_TeXToggle"); }
                catch { throw new Exception("MathType 转换命令不可用。请确认 Microsoft Word 顶部有 MathType 选项卡。"); }
                if ((int)doc.InlineShapes.Count != 1) throw new Exception("MathType 未能转换这条公式，已保留原代码。");
                string prog = (string)doc.InlineShapes[1].OLEFormat.ProgID;
                if (prog.IndexOf("DSMT", StringComparison.OrdinalIgnoreCase) < 0 && prog.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) < 0) throw new Exception("结果不是 MathType 对象，已保留原代码。");
                if (formula.Number.Length > 0) { dynamic r = doc.Range((int)doc.InlineShapes[1].Range.End, (int)doc.InlineShapes[1].Range.End); r.InsertAfter("  (" + formula.Number + ")"); }
                doc.SaveAs2(output, 6); CloseDocument(doc); doc = null;
                return new FormulaRender { Rtf = File.ReadAllText(output, Encoding.Default), ProgId = prog };
            }
            finally { CloseDocument(doc); Remove(output); }
        }
        internal string HtmlToRtf(string html)
        {
            Ensure(); dynamic doc = null;
            string source = NewPath(".html"), output = NewPath(".rtf");
            try
            {
                File.WriteAllText(source, HtmlMath.Prepare(html), new UTF8Encoding(true));
                doc = word.Documents.Open(source, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                doc.SaveAs2(output, 6); CloseDocument(doc); doc = null;
                return File.ReadAllText(output, Encoding.Default);
            }
            finally { CloseDocument(doc); Remove(source); Remove(output); }
        }
        internal string VerifyRtf(string rtf, int expectedObjects)
        {
            Ensure(); dynamic doc = null; string source = NewPath(".rtf");
            try
            {
                File.WriteAllText(source, rtf, Encoding.ASCII);
                doc = word.Documents.Open(source, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                int count = 0; foreach (dynamic s in doc.InlineShapes) try { string p = (string)s.OLEFormat.ProgID; if (p.Contains("DSMT")) count++; } catch { }
                if (count != expectedObjects) throw new Exception("富文本中的 MathType 对象数量不符：" + count + "/" + expectedObjects);
                return (string)doc.Content.Text;
            }
            finally { CloseDocument(doc); Remove(source); }
        }
        internal TextStyle InspectStyle(string rtf, string text)
        {
            Ensure(); dynamic doc = null; string source = NewPath(".rtf");
            try
            {
                File.WriteAllText(source, rtf, Encoding.ASCII);
                doc = word.Documents.Open(source, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                string all = (string)doc.Content.Text; int index = all.IndexOf(text, StringComparison.Ordinal);
                if (index < 0) throw new Exception("测试文字未找到。");
                dynamic f = doc.Range(index, index + text.Length).Font;
                return new TextStyle { Font = (string)f.Name, Size = (float)f.Size, Bold = (int)f.Bold, Color = (int)f.Color };
            }
            finally { CloseDocument(doc); Remove(source); }
        }
        internal string Insert(string rtf, IntPtr target)
        {
            dynamic win = Office.WindowObject(target), app = win.Application;
            win.Activate(); dynamic doc = win.Document;
            if ((bool)doc.ReadOnly || (int)doc.ProtectionType != -1) throw new Exception("目标 Word 文档只读或受保护。编辑区内容已保留。");
            dynamic selected = win.Selection;
            if ((int)selected.StoryType != 1) throw new Exception("请把 Word 光标放在正文中，再插入。");
            dynamic range = selected.Range.Duplicate;
            // Insert after any selection. Never overwrite the selected source content.
            range.Collapse(0); int start = (int)range.Start; int oldEnd = (int)doc.Content.End;
            string source = NewPath(".rtf"); File.WriteAllText(source, rtf, Encoding.ASCII);
            bool undo = false;
            try
            {
                try { app.UndoRecord.StartCustomRecord("从 MathType 编辑器插入"); undo = true; } catch { }
                range.InsertFile(source, ConfirmConversions: false, Link: false, Attachment: false);
                int end = start + Math.Max(0, (int)doc.Content.End - oldEnd); selected.SetRange(end, end);
                app.Activate(); Office.SetForegroundWindow(target);
                return "已插入“" + doc.Name + "”。正文和公式一起保留；Word 中可按 Ctrl+Z 撤销。";
            }
            finally { if (undo) try { app.UndoRecord.EndCustomRecord(); } catch { } Remove(source); }
        }
        public void Dispose()
        {
            if (word != null) { try { word.Quit(0); } catch { } word = null; }
            if (process != null) { try { if (!process.WaitForExit(1500) && !process.HasExited) process.Kill(); } catch { } process.Dispose(); process = null; }
            Office.PrivatePid = 0;
            // Remove only files created in this unique session, never a caller-supplied directory.
            try { foreach (string file in Directory.GetFiles(session)) Remove(file); Directory.Delete(session, false); } catch { }
        }
    }
    internal sealed class WordWorker : IDisposable
    {
        readonly BlockingCollection<Action<WordEngine>> jobs = new BlockingCollection<Action<WordEngine>>();
        readonly Thread thread;
        internal WordWorker()
        {
            thread = new Thread(delegate()
            {
                using (var engine = new WordEngine())
                {
                    while (!jobs.IsCompleted)
                    {
                        Action<WordEngine> job;
                        if (jobs.TryTake(out job, 80)) job(engine);
                        Application.DoEvents();
                    }
                }
            });
            thread.SetApartmentState(ApartmentState.STA); thread.IsBackground = true; thread.Start();
        }
        internal Task<T> Run<T>(Func<WordEngine, T> action)
        {
            var done = new TaskCompletionSource<T>();
            jobs.Add(delegate(WordEngine engine) { try { done.SetResult(action(engine)); } catch (Exception e) { done.SetException(e); } });
            return done.Task;
        }
        public void Dispose() { jobs.CompleteAdding(); thread.Join(5000); }
    }
}
