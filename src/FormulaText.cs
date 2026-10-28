using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MathTypeEditor
{
    internal sealed class FormulaSpec
    {
        internal int Start, Length;
        internal string Body, Number;
        internal bool Display;
    }
    internal static class FormulaText
    {
        static readonly Regex Blocks = new Regex(
            @"```(?:latex|tex|math)[ \t]*\r?\n[\s\S]*?\r?\n```" +
            @"|\\begin\{equation\*?\}[\s\S]*?\\end\{equation\*?\}" +
            @"|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\)" +
            @"|(?<![\\$])\$\$[\s\S]*?\$\$(?!\$)" +
            @"|(?<![\\$])\$[^\r\n$]+?(?<!\\)\$(?!\$)" +
            @"|(?m:^[ \t]*\[[ \t]*(?:\r?\n|\v)[\s\S]*?(?:\r?\n|\v)[ \t]*\][ \t]*(?=\r?$))" +
            @"|(?<!\\)\((?=[^()\r\n]*\\(?:frac|times|mathrm|sqrt|alpha|beta|sum|int)\b)[^()\r\n]+\)",
            RegexOptions.Compiled);
        internal static List<FormulaSpec> Scan(string text, out List<string> warnings)
        {
            var found = new List<FormulaSpec>(); warnings = new List<string>();
            foreach (Match m in Blocks.Matches(text))
            {
                bool display = m.Value.StartsWith(@"\[") || m.Value.StartsWith("$$") || m.Value.StartsWith(@"\begin") || m.Value.StartsWith("```") || m.Value.TrimStart().StartsWith("[");
                if (m.Value.StartsWith("$") && !m.Value.StartsWith("$$") && Regex.IsMatch(Regex.Replace(m.Value, @"\\(?:text|mathrm)\{[^{}]*\}", ""), @"[\u4e00-\u9fff]")) continue;
                if (m.Value.TrimStart().StartsWith("[") && !m.Value.Contains(@"\") && !Regex.IsMatch(m.Value, @"[=^_]")) continue;
                try
                {
                    FormulaSpec f = Parse(m.Value, display, m.Value.StartsWith("(")); f.Start = m.Index; f.Length = m.Length; found.Add(f);
                }
                catch (Exception e) { warnings.Add("第 " + (m.Index + 1) + " 字附近：" + e.Message); }
            }
            return found;
        }
        internal static FormulaSpec Parse(string source, bool display, bool stripPlainParentheses = false)
        {
            string s = source.Trim().Trim('\uFEFF', '\u200B').Trim();
            Match fence = Regex.Match(s, @"\A```(?:latex|tex|math)[ \t]*\r?\n([\s\S]*?)\r?\n```\z");
            if (fence.Success) s = fence.Groups[1].Value.Trim();
            var pairs = new[] { new[] { @"\[", @"\]" }, new[] { @"\(", @"\)" }, new[] { "$$", "$$" }, new[] { "$", "$" }, new[] { @"\begin{equation*}", @"\end{equation*}" }, new[] { @"\begin{equation}", @"\end{equation}" } };
            bool again = true;
            while (again)
            {
                again = false;
                foreach (var pair in pairs) if (s.StartsWith(pair[0]) && s.EndsWith(pair[1]) && s.Length > pair[0].Length + pair[1].Length)
                { s = s.Substring(pair[0].Length, s.Length - pair[0].Length - pair[1].Length).Trim(); again = true; break; }
            }
            s = Regex.Replace(s, @"\A\[[ \t]*(?:\r?\n|\v)([\s\S]*?)(?:\r?\n|\v)[ \t]*\]\z", "$1").Trim();
            if (stripPlainParentheses && s.StartsWith("(") && s.EndsWith(")")) s = s.Substring(1, s.Length - 2).Trim();
            var tags = Regex.Matches(s, @"\\tag\*?\s*\{([^{}]*)\}");
            if (tags.Count > 1) throw new Exception("一条公式中包含多个 tag 编号。");
            string number = tags.Count == 1 ? tags[0].Groups[1].Value.Trim().Trim('(', ')', '（', '）') : "";
            if (number.Length > 50 || Regex.IsMatch(number, @"[\\{}\r\n]")) throw new Exception("编号不受支持。");
            s = Regex.Replace(s, @"\\(?:tag\*?|label)\s*\{[^{}]*\}", "");
            s = Regex.Replace(s, @"(?<!\\)%[^\r\n]*", "");
            if (Regex.IsMatch(s, @"\\[\[\]]|\$|```|\\(?:documentclass|usepackage|input|include|write)\b")) throw new Exception("公式边界不完整，或包含不支持的命令。");
            int depth = 0;
            for (int i = 0; i < s.Length; i++)
            {
                int slash = 0; for (int j = i - 1; j >= 0 && s[j] == '\\'; j--) slash++;
                if (slash % 2 != 0) continue;
                if (s[i] == '{') depth++;
                if (s[i] == '}' && --depth < 0) throw new Exception("花括号不成对。");
            }
            if (depth != 0) throw new Exception("花括号不成对。");
            s = Regex.Replace(s, @"\s*\r?\n\s*", " ").Trim();
            if (s.Length == 0 || s.Length > 20000) throw new Exception("公式为空或过长。");
            return new FormulaSpec { Body = s, Number = number, Display = display };
        }
    }

    internal static class HtmlMath
    {
        sealed class Node
        {
            internal string Tag, Open;
            internal int Start, OpenEnd, End, CloseStart;
            internal Node Parent;
        }
        sealed class Replacement { internal int Start, End; internal string Text; }
        static string Attribute(string open, string attr)
        {
            Match m = Regex.Match(open, @"\b" + Regex.Escape(attr) + @"\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]+))", RegexOptions.IgnoreCase);
            return m.Success ? WebUtility.HtmlDecode(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value) : "";
        }
        internal static string FromClipboard(string content)
        {
            Match a = Regex.Match(content, @"StartHTML:(\d+)"), b = Regex.Match(content, @"EndHTML:(\d+)");
            if (a.Success && b.Success)
            {
                byte[] data = Encoding.UTF8.GetBytes(content); int start, end;
                if (Int32.TryParse(a.Groups[1].Value, out start) && Int32.TryParse(b.Groups[1].Value, out end) && start >= 0 && end > start && end <= data.Length) return Encoding.UTF8.GetString(data, start, end - start);
            }
            int html = content.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
            if (html >= 0) return content.Substring(html);
            int fragment = content.IndexOf("<!--StartFragment-->", StringComparison.OrdinalIgnoreCase);
            return fragment >= 0 ? content.Substring(fragment) : content;
        }
        internal static string Prepare(string content)
        {
            string html = FromClipboard(content);
            html = Regex.Replace(html, @"<script\b([^>]*\btype\s*=\s*['""]math/tex[^>]*?)>([\s\S]*?)</script\s*>", delegate(Match m)
            { bool block = m.Groups[1].Value.IndexOf("display", StringComparison.OrdinalIgnoreCase) >= 0; return Wrap(m.Groups[2].Value, block); }, RegexOptions.IgnoreCase);
            var nodes = new List<Node>(); var stack = new List<Node>();
            foreach (Match m in Regex.Matches(html, @"<!--[^>]*-->|<(/?)([A-Za-z][\w:-]*)(?:[^>""']|""[^""]*""|'[^']*')*>"))
            {
                if (m.Value.StartsWith("<!--")) continue;
                string tag = m.Groups[2].Value.ToLowerInvariant(); bool closing = m.Groups[1].Value.Length > 0;
                if (closing)
                {
                    for (int j = stack.Count - 1; j >= 0; j--) if (stack[j].Tag == tag)
                    { stack[j].End = m.Index + m.Length; stack[j].CloseStart = m.Index; stack.RemoveRange(j, stack.Count - j); break; }
                    continue;
                }
                var n = new Node { Tag = tag, Open = m.Value, Start = m.Index, OpenEnd = m.Index + m.Length, End = m.Index + m.Length, CloseStart = m.Index + m.Length, Parent = stack.Count == 0 ? null : stack[stack.Count - 1] };
                nodes.Add(n);
                if (!m.Value.EndsWith("/>") && !new[] { "img", "br", "hr", "meta", "link", "input", "wbr", "source" }.Contains(tag)) stack.Add(n);
            }
            var changes = new List<Replacement>();
            foreach (Node n in nodes)
            {
                string tex = ""; bool display = false; Node root = n;
                if (n.Tag == "annotation" && Attribute(n.Open, "encoding").IndexOf("tex", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    tex = WebUtility.HtmlDecode(html.Substring(n.OpenEnd, n.CloseStart - n.OpenEnd));
                    for (Node p = n.Parent; p != null; p = p.Parent)
                    {
                        string cls = Attribute(p.Open, "class");
                        if (p.Tag == "math" || p.Tag == "mjx-container" || Regex.IsMatch(cls, @"(?:^|\s)(?:katex|katex-display|MathJax|MathJax_Display)(?:\s|$)")) root = p;
                        if (cls.Contains("katex-display") || cls.Contains("MathJax_Display") || Attribute(p.Open, "display") == "block" || Attribute(p.Open, "display") == "true") display = true;
                    }
                }
                else
                {
                    tex = Attribute(n.Open, "data-latex"); if (tex.Length == 0) tex = Attribute(n.Open, "data-tex");
                    if (tex.Length == 0 && n.Tag == "img") { string alt = Attribute(n.Open, "alt"); if (alt.StartsWith(@"\") || alt.StartsWith("$")) tex = alt; }
                    if (tex.Length == 0 && n.Tag == "math" && !nodes.Any(a => a.Tag == "annotation" && a.Start > n.Start && a.End <= n.End && Attribute(a.Open, "encoding").IndexOf("tex", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        try { tex = MathMl(html.Substring(n.Start, n.End - n.Start)); } catch { tex = ""; }
                        display = Attribute(n.Open, "display") == "block";
                    }
                }
                if (String.IsNullOrWhiteSpace(tex)) continue;
                if (tex.StartsWith(@"\[") || tex.StartsWith("$$")) display = true;
                try { tex = FormulaText.Parse(tex, display).Body + ExtractTag(tex); } catch { }
                changes.Add(new Replacement { Start = root.Start, End = root.End, Text = Wrap(tex, display) });
            }
            // Pick the outer representation once: KaTeX commonly contains both MathML and visual spans.
            var chosen = new List<Replacement>();
            foreach (var r in changes.OrderBy(r => r.Start).ThenByDescending(r => r.End))
                if (!chosen.Any(c => r.Start >= c.Start && r.Start < c.End)) chosen.Add(r);
            foreach (var r in chosen.OrderByDescending(r => r.Start)) html = html.Substring(0, r.Start) + r.Text + html.Substring(r.End);
            html = Regex.Replace(html, @"<(script|iframe|object|embed)\b[^>]*>[\s\S]*?</\1\s*>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"\s+on\w+\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"<meta\b[^>]*charset[^>]*>", "", RegexOptions.IgnoreCase);
            if (Regex.IsMatch(html, @"<html\b", RegexOptions.IgnoreCase))
            {
                var head = new Regex(@"<head\b[^>]*>", RegexOptions.IgnoreCase);
                if (head.IsMatch(html)) return head.Replace(html, "$0<meta charset=\"utf-8\">", 1);
                return new Regex(@"<html\b[^>]*>", RegexOptions.IgnoreCase).Replace(html, "$0<head><meta charset=\"utf-8\"></head>", 1);
            }
            return "<html><head><meta charset=\"utf-8\"></head><body>" + html + "</body></html>";
        }
        static string ExtractTag(string source)
        { Match m = Regex.Match(source, @"\\tag\*?\s*\{[^{}]*\}"); return m.Success ? m.Value : ""; }
        static string Wrap(string tex, bool display)
        { string encoded = WebUtility.HtmlEncode((display ? @"\[" : @"\(") + tex + (display ? @"\]" : @"\)")); return display ? "<div>" + encoded + "</div>" : "<span>" + encoded + "</span>"; }
        static string MathMl(string source)
        {
            source = Regex.Replace(source, @"&[A-Za-z][A-Za-z0-9]*;", delegate(Match m) { string decoded = WebUtility.HtmlDecode(m.Value); if (decoded == m.Value) return m.Value; return String.Concat(decoded.Select(c => "&#" + ((int)c) + ";")); });
            return Ml(XElement.Parse(source));
        }
        static string Ml(XElement e)
        {
            var children = e.Elements().ToArray(); string t = e.Name.LocalName.ToLowerInvariant();
            Func<int, string> c = i => i < children.Length ? Ml(children[i]) : "";
            switch (t)
            {
                case "annotation": return "";
                case "semantics": return c(0);
                case "mi": case "mn": return e.Value;
                case "mtext": return @"\text{" + e.Value.Replace("{", @"\{").Replace("}", @"\}") + "}";
                case "mo":
                    switch (e.Value) { case "×": return @"\times "; case "≤": return @"\leq "; case "≥": return @"\geq "; case "≈": return @"\approx "; case "−": return "-"; case "∞": return @"\infty "; case "∑": return @"\sum "; case "∫": return @"\int "; case "\u2062": case "\u2061": return ""; default: return e.Value; }
                case "msup": return "{" + c(0) + "}^{" + c(1) + "}";
                case "msub": return "{" + c(0) + "}_{" + c(1) + "}";
                case "msubsup": return "{" + c(0) + "}_{" + c(1) + "}^{" + c(2) + "}";
                case "mfrac": return @"\frac{" + c(0) + "}{" + c(1) + "}";
                case "msqrt": return @"\sqrt{" + String.Concat(children.Select(Ml)) + "}";
                case "mroot": return @"\sqrt[" + c(1) + "]{" + c(0) + "}";
                case "mfenced": return @"\left" + ((string)e.Attribute("open") ?? "(") + String.Join(",", children.Select(Ml)) + @"\right" + ((string)e.Attribute("close") ?? ")");
                case "mtable": return @"\begin{matrix}" + String.Join(@"\\", children.Select(Ml)) + @"\end{matrix}";
                case "mtr": return String.Join("&", children.Select(Ml));
                case "mover": if (children.Length > 1 && (children[1].Value == "^" || children[1].Value == "ˆ")) return @"\hat{" + c(0) + "}"; return @"\overset{" + c(1) + "}{" + c(0) + "}";
                case "munder": return @"\underset{" + c(1) + "}{" + c(0) + "}";
                case "munderover": return "{" + c(0) + "}_{" + c(1) + "}^{" + c(2) + "}";
                default: return children.Length > 0 ? String.Concat(children.Select(Ml)) : e.Value;
            }
        }
    }
}
