using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    /// <summary>A source-preserving expression. Offsets are zero based in the original formula, including its initial '='.</summary>
    public sealed class ExpressionNode
    {
        public string Text, Kind;
        public int Start, Length;
        public readonly List<ExpressionNode> Children = new List<ExpressionNode>();
        internal SheetSpanReference SheetSpan;
        internal ExpressionNode SheetSpanOrigin;
        public override string ToString() { return Text; }
    }

    internal sealed class SheetSpanReference
    {
        static readonly Regex Qualifier = new Regex(@"\G(?<s>'(?:[^']|'')+'|(?:\[[^\]]+\])?[\p{L}\p{N}_.]+:[\p{L}\p{N}_.]+)!");
        internal readonly string First, Last, Address, Workbook;
        internal readonly bool HasPath;
        internal ExpressionNode[] Parts = new ExpressionNode[0];

        SheetSpanReference(string first, string last, string address, string workbook, bool hasPath)
        { First = first; Last = last; Address = address; Workbook = workbook; HasPath = hasPath; }

        internal IEnumerable<string> WorksheetsBetween(IList<string> order)
        {
            int first = -1;
            for (int i = 0; i < order.Count; i++)
            {
                if (first < 0 && String.Equals(order[i], First, StringComparison.OrdinalIgnoreCase)) first = i;
                if (!String.Equals(order[i], Last, StringComparison.OrdinalIgnoreCase)) continue;
                if (first >= 0)
                    for (int sheet = first; sheet <= i; sheet++) yield return order[sheet];
                yield break;
            }
        }

        internal static SheetSpanReference Read(string source, int start, out int end)
        {
            end = start;
            Match qualifier = Qualifier.Match(source, start);
            if (!qualifier.Success) return null;
            string sheets = qualifier.Groups["s"].Value;
            if (sheets[0] == '\'') sheets = sheets.Substring(1, sheets.Length - 2).Replace("''", "'");
            int bracket = sheets.LastIndexOf(']');
            int bookStart = bracket < 0 ? -1 : sheets.LastIndexOf('[', bracket);
            if (bracket >= 0 && (bookStart < 0 || bracket == bookStart + 1)) return null;
            string workbook = bookStart < 0 ? "" : sheets.Substring(bookStart + 1, bracket - bookStart - 1);
            string[] names = sheets.Substring(bracket + 1).Split(':');
            if (names.Length != 2 || names[0].Length == 0 || names[1].Length == 0) return null;
            Match address = FormulaParser.AddressPattern.Match(source, start + qualifier.Length);
            if (!address.Success || Reference.Parse(names[0], address.Value) == null) return null;
            end = address.Index + address.Length;
            return new SheetSpanReference(names[0], names[1], address.Value, workbook, bookStart > 0);
        }
    }

    /// <summary>
    /// Parses Excel's invariant Formula/Formula2 syntax, without calculating values or resolving names.
    /// Local decimal commas, R1C1 references and DDE expressions are not supported. Function argument
    /// separators may be commas or semicolons; array columns/rows use invariant commas/semicolons.
    /// </summary>
    public static class FormulaSyntax
    {
        public static ExpressionNode Parse(string formula)
        {
            if (formula == null) throw new ArgumentNullException("formula");
            if (formula.Length > 32768) throw new FormatException("Formula exceeds the supported 32768-character limit.");
            return new Parser(formula).Parse();
        }

        public static IEnumerable<ExpressionNode> Flatten(ExpressionNode root)
        {
            if (root == null) yield break;
            var pending = new Stack<ExpressionNode>();
            pending.Push(root);
            while (pending.Count != 0)
            {
                ExpressionNode node = pending.Pop();
                yield return node;
                for (int i = node.Children.Count - 1; i >= 0; i--) pending.Push(node.Children[i]);
            }
        }

        /// <summary>Returns an invariant expression with enclosing LET bindings needed to evaluate a selected subexpression.</summary>
        public static string EvaluationText(ExpressionNode root, ExpressionNode target)
        { return EvaluationText(root, target, null); }

        internal static string EvaluationText(ExpressionNode root, ExpressionNode target, string replacement)
        {
            if (root == null || target == null) throw new ArgumentNullException(root == null ? "root" : "target");
            if (target.SheetSpanOrigin != null)
            {
                replacement = replacement ?? target.Text;
                target = target.SheetSpanOrigin;
            }
            var path = new List<ExpressionNode>();
            if (!FindPath(root, target, path)) throw new ArgumentException("The selected expression does not belong to this formula.", "target");
            string expression = replacement ?? target.Text;
            if (target.Kind == "Missing argument" || target.Kind == "Array row") throw new InvalidOperationException("This syntax element is not independently evaluable.");
            for (int p = path.Count - 2; p >= 0; p--)
            {
                ExpressionNode scope = path[p];
                string name = FunctionName(scope);
                if (name == "LAMBDA") throw new InvalidOperationException("A LAMBDA subexpression cannot be evaluated without its bound invocation parameters.");
                if (name != "LET") continue;
                int child = scope.Children.IndexOf(path[p + 1]);
                if (scope.Children.Count < 3 || scope.Children.Count % 2 != 1)
                    throw new InvalidOperationException("LET requires name/value pairs and a final calculation.");
                if (child < scope.Children.Count - 1 && child % 2 == 0)
                    throw new InvalidOperationException("A LET binding name is a declaration, not an independently evaluable expression.");
                int visiblePairs = child / 2;
                if (visiblePairs == 0) continue;
                var prefix = new List<string>();
                for (int i = 0; i < visiblePairs * 2; i++) prefix.Add(scope.Children[i].Text);
                expression = "LET(" + String.Join(",", prefix.ToArray()) + "," + expression + ")";
            }
            return expression;
        }

        internal static string FunctionName(ExpressionNode node)
        {
            if (node.Kind != "Function") return "";
            int parenthesis = node.Text.IndexOf('(');
            string name = parenthesis < 0 ? "" : node.Text.Substring(0, parenthesis).Trim().ToUpperInvariant();
            if (name.StartsWith("_XLFN.", StringComparison.Ordinal)) name = name.Substring(6);
            return name.StartsWith("_XLWS.", StringComparison.Ordinal) ? name.Substring(6) : name;
        }

        internal static ExpressionNode UnwrapGroups(ExpressionNode node)
        {
            while (node.Kind == "Group" && node.Children.Count == 1) node = node.Children[0];
            return node;
        }

        internal static ExpressionNode LetBinding(ExpressionNode root, ExpressionNode variable)
        {
            var path = new List<ExpressionNode>();
            if (!FindPath(root, variable, path)) return null;
            for (int p = path.Count - 2; p >= 0; p--)
            {
                ExpressionNode scope = path[p];
                if (FunctionName(scope) != "LET" || scope.Children.Count < 3 || scope.Children.Count % 2 == 0) continue;
                int visiblePairs = scope.Children.IndexOf(path[p + 1]) / 2;
                for (int i = visiblePairs * 2 - 2; i >= 0; i -= 2)
                    if (String.Equals(scope.Children[i].Text, variable.Text, StringComparison.OrdinalIgnoreCase)) return scope.Children[i + 1];
            }
            return null;
        }

        internal static bool FindPath(ExpressionNode node, ExpressionNode target, List<ExpressionNode> path)
        {
            var parents = new Dictionary<ExpressionNode, ExpressionNode>();
            var pending = new Stack<ExpressionNode>();
            parents[node] = null;
            pending.Push(node);
            while (pending.Count != 0)
            {
                ExpressionNode current = pending.Pop();
                if (Object.ReferenceEquals(current, target))
                {
                    while (current != null) { path.Add(current); current = parents[current]; }
                    path.Reverse(); return true;
                }
                foreach (ExpressionNode child in current.Children)
                {
                    if (parents.ContainsKey(child)) continue;
                    parents[child] = current; pending.Push(child);
                }
            }
            return false;
        }

        internal static IReadOnlyList<Token> PresentationTokens(string formula, FormulaDisplayStyle style, Regex gridReference)
        { return new Parser(formula, true, style, gridReference).Tokens; }

        internal static bool IsInvariantFormulaError(string text)
        { return Array.Exists(Parser.Errors, error => String.Equals(error, text, StringComparison.OrdinalIgnoreCase)); }

        internal sealed class Token
        {
            public string Text, Kind;
            public int Start, End;
            public SheetSpanReference SheetSpan;
        }

        sealed class Parser
        {
            readonly string source;
            readonly List<Token> tokens = new List<Token>();
            readonly bool presentation;
            readonly FormulaDisplayStyle displayStyle;
            readonly Regex displayGridReference;
            int index, depth, arrayDepth;
            static readonly Regex R1C1Token = new Regex(@"^R(?:\d+|\[-?\d+\])?C(?:\d+|\[-?\d+\])?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            internal static readonly string[] Errors = { "#GETTING_DATA", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#N/A", "#SPILL!", "#CALC!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#BUSY!", "#PYTHON!" };

            public Parser(string text, bool presentation = false, FormulaDisplayStyle displayStyle = null, Regex displayGridReference = null)
            {
                source = text;
                this.presentation = presentation;
                this.displayStyle = displayStyle;
                this.displayGridReference = displayGridReference;
                Lex();
            }
            internal IReadOnlyList<Token> Tokens { get { return tokens; } }
            Token Current { get { return tokens[index]; } }
            Token Take() { return tokens[index++]; }
            FormatException Error(string message, int position) { return new FormatException(message + " (position " + position + ")."); }
            void Expect(string value)
            {
                if (Current.Text != value) throw Error("Expected '" + value + "'", Current.Start);
                Take();
            }
            ExpressionNode Node(string kind, int start, int end, params ExpressionNode[] children)
            {
                var result = new ExpressionNode { Kind = kind, Start = start, Length = end - start, Text = source.Substring(start, end - start) };
                result.Children.AddRange(children);
                return result;
            }

            public ExpressionNode Parse()
            {
                if (Current.Text == "=") Take();
                if (Current.Kind == "End") throw Error("Empty formula", Current.Start);
                ExpressionNode root = Expression(0, true);
                if (Current.Kind != "End") throw Error("Unexpected token '" + Current.Text + "'", Current.Start);
                return root;
            }

            ExpressionNode Expression(int minimum, bool union)
            {
                if (++depth > 256) throw Error("Formula nesting exceeds 256 levels", Current.Start);
                try
                {
                    ExpressionNode left = Prefix(union);
                    while (true)
                    {
                        Token token = Current;
                        if (token.Text == "(" && token.Start == left.Start + left.Length && (left.Kind == "Function" || left.Kind == "Invocation" || left.Kind == "Group"))
                        {
                            if (110 < minimum) break;
                            left = Call(left, true); continue;
                        }
                        if (token.Text == "%" || token.Text == "#")
                        {
                            int postfix = token.Text == "#" ? 110 : 60;
                            if (postfix < minimum) break;
                            if (token.Text == "#" && !ReferenceLike(left)) throw Error("Spill operator requires a reference expression", token.Start);
                            Take();
                            left = Node(token.Text == "#" ? "Spill" : "Percent", left.Start, token.End, left);
                            continue;
                        }
                        bool intersection = token.Start > left.Start + left.Length && ReferenceLike(left) && StartsReference(token);
                        string operation = intersection ? " " : token.Text;
                        int precedence = Precedence(operation, union);
                        if (precedence < minimum || precedence < 0) break;
                        if (!intersection) Take();
                        ExpressionNode right = Expression(precedence + 1, union);
                        if ((operation == "," || operation == " ") && (!ReferenceLike(left) || !ReferenceLike(right)))
                            throw Error("Reference operator requires reference expressions", token.Start);
                        if (operation == ":" && (!RangeEndpoint(left) || !RangeEndpoint(right)))
                            throw Error("Range operator requires references or row numbers", token.Start);
                        left = Node(operation == ":" ? "Range" : operation == "," ? "Union" : operation == " " ? "Intersection" : "Operator " + operation,
                            left.Start, right.Start + right.Length, left, right);
                    }
                    return left;
                }
                finally { depth--; }
            }

            static int Precedence(string value, bool union)
            {
                switch (value)
                {
                    case ":": return 100;
                    case " ": return 90;
                    case ",": return union ? 80 : -1;
                    // Excel applies negation before percent and exponentiation, and equal-precedence operators left to right.
                    case "^": return 50;
                    case "*": case "/": return 40;
                    case "+": case "-": return 30;
                    case "&": return 20;
                    case "=": case "<>": case "<": case ">": case "<=": case ">=": return 10;
                    default: return -1;
                }
            }

            static bool ReferenceLike(ExpressionNode node)
            {
                if (node.Kind == "Group") return node.Children.Count == 1 && ReferenceLike(node.Children[0]);
                return node.SheetSpan != null || node.Kind == "Reference" || node.Kind == "Name" || node.Kind == "Function" || node.Kind == "Invocation" || node.Kind == "Range" || node.Kind == "Union" || node.Kind == "Intersection" || node.Kind == "Spill";
            }
            static bool RangeEndpoint(ExpressionNode node)
            {
                return ReferenceLike(node) || (node.Kind == "Number" && Regex.IsMatch(node.Text, @"^[1-9][0-9]*$"));
            }
            static bool StartsReference(Token token) { return token.Kind == "Atom" || token.Text == "("; }

            ExpressionNode Prefix(bool union)
            {
                Token token = Take();
                if (token.SheetSpan != null)
                {
                    ExpressionNode span = Node("3D reference", token.Start, token.End);
                    span.SheetSpan = token.SheetSpan;
                    return span;
                }
                if (token.Text == "+" || token.Text == "-" || token.Text == "@")
                {
                    ExpressionNode operand = Expression(70, union);
                    return Node(token.Text == "@" ? "Implicit intersection" : "Unary " + token.Text, token.Start, operand.Start + operand.Length, operand);
                }
                if (token.Text == "(")
                {
                    ExpressionNode operand = Expression(0, true);
                    int end = Current.End;
                    Expect(")");
                    return Node("Group", token.Start, end, operand);
                }
                if (token.Text == "{") return Array(token);
                if (token.Kind == "Number" || token.Kind == "String" || token.Kind == "Error") return Node(token.Kind, token.Start, token.End);
                if (token.Kind == "Atom")
                {
                    if (R1C1Token.IsMatch(token.Text) && !String.Equals(token.Text, "RC", StringComparison.OrdinalIgnoreCase))
                        throw Error("R1C1 references are unsupported; use Formula or Formula2 in A1 notation", token.Start);
                    bool boolean = String.Equals(token.Text, "TRUE", StringComparison.OrdinalIgnoreCase) || String.Equals(token.Text, "FALSE", StringComparison.OrdinalIgnoreCase);
                    int row, column;
                    bool validCell = Coordinates.TryCell(token.Text, out row, out column);
                    bool reference = token.Text.IndexOfAny(new[] { '!', '[', '$' }) >= 0 || validCell;
                    ExpressionNode atom = Node(boolean ? "Boolean" : reference ? "Reference" : "Name", token.Start, token.End);
                    if (Current.Text == "(" && !boolean && !(reference && Current.Start > token.End && !String.Equals(token.Text, "LOG10", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (validCell && !String.Equals(token.Text, "LOG10", StringComparison.OrdinalIgnoreCase))
                            throw Error("A cell address cannot be used as a function name", token.Start);
                        if (reference && token.Text.IndexOf('!') < 0 && !String.Equals(token.Text, "LOG10", StringComparison.OrdinalIgnoreCase))
                            throw Error("Invalid function name", token.Start);
                        return Call(atom, false);
                    }
                    return atom;
                }
                throw Error("Expected an expression; found '" + token.Text + "'", token.Start);
            }

            ExpressionNode Call(ExpressionNode target, bool invocation)
            {
                Expect("(");
                var arguments = new List<ExpressionNode>();
                if (invocation) arguments.Add(target);
                string separator = null;
                bool expectingArgument = true, seenSeparator = false;
                while (Current.Text != ")")
                {
                    if (Current.Kind == "End") throw Error("Unclosed function call", Current.Start);
                    if (Current.Text == "," || Current.Text == ";")
                    {
                        if (separator != null && separator != Current.Text) throw Error("Mixed function argument separators", Current.Start);
                        separator = Current.Text;
                        if (expectingArgument) arguments.Add(Node("Missing argument", Current.Start, Current.Start));
                        Take(); expectingArgument = true; seenSeparator = true;
                    }
                    else
                    {
                        if (!expectingArgument) throw Error("Expected an argument separator", Current.Start);
                        arguments.Add(Expression(0, false));
                        expectingArgument = false;
                    }
                }
                if (expectingArgument && seenSeparator) arguments.Add(Node("Missing argument", Current.Start, Current.Start));
                Token closing = Take();
                return Node(invocation ? "Invocation" : "Function", target.Start, closing.End, arguments.ToArray());
            }

            ExpressionNode Array(Token opening)
            {
                var rows = new List<ExpressionNode>();
                var values = new List<ExpressionNode>();
                int rowStart = Current.Start, width = -1;
                while (true)
                {
                    if (Current.Kind == "End") throw Error("Unclosed array constant", Current.Start);
                    ExpressionNode value = Expression(0, false);
                    if (!ArrayScalar(value)) throw Error("Array constants contain only numbers, strings, booleans and errors", value.Start);
                    values.Add(value);
                    if (Current.Text == ",") { Take(); continue; }
                    if (Current.Text != ";" && Current.Text != "}") throw Error("Expected an array separator or '}'", Current.Start);
                    if (width < 0) width = values.Count;
                    if (values.Count != width) throw Error("Array constant rows have different widths", rowStart);
                    rows.Add(Node("Array row", rowStart, value.Start + value.Length, values.ToArray()));
                    values.Clear();
                    Token separator = Take();
                    if (separator.Text == "}") return Node("Array", opening.Start, separator.End, rows.ToArray());
                    rowStart = Current.Start;
                }
            }
            static bool ArrayScalar(ExpressionNode node)
            {
                if (node.Kind == "Number" || node.Kind == "String" || node.Kind == "Boolean" || node.Kind == "Error") return true;
                return (node.Kind == "Unary -" || node.Kind == "Unary +") && node.Children.Count == 1 && node.Children[0].Kind == "Number";
            }

            void Lex()
            {
                int p = 0;
                char decimalPoint = displayStyle == null ? '.' : displayStyle.DecimalSeparator;
                while (p < source.Length)
                {
                    if (Char.IsWhiteSpace(source[p])) { p++; continue; }
                    int start = p;
                    char c = source[p];
                    if (!presentation && (c == '\'' || c == '[' || IsNameCharacter(c)))
                    {
                        int end;
                        SheetSpanReference span = SheetSpanReference.Read(source, p, out end);
                        if (span != null)
                        {
                            tokens.Add(new Token { Start = start, End = end, Kind = "Atom", Text = source.Substring(start, end - start), SheetSpan = span });
                            p = end;
                            continue;
                        }
                    }
                    string kind = "Symbol";
                    if (c == '"')
                    {
                        kind = "String"; p++; bool closed = false;
                        while (p < source.Length)
                        {
                            if (source[p++] != '"') continue;
                            if (p < source.Length && source[p] == '"') { p++; continue; }
                            closed = true; break;
                        }
                        if (!closed) throw Error("Unclosed string literal", start);
                    }
                    else if (Char.IsDigit(c) || (c == decimalPoint && p + 1 < source.Length && Char.IsDigit(source[p + 1])))
                    {
                        kind = "Number";
                        while (p < source.Length && Char.IsDigit(source[p])) p++;
                        if (p < source.Length && source[p] == decimalPoint) { p++; while (p < source.Length && Char.IsDigit(source[p])) p++; }
                        if (p + 1 < source.Length && (source[p] == 'E' || source[p] == 'e') && (Char.IsDigit(source[p + 1]) || source[p + 1] == '+' || source[p + 1] == '-'))
                        {
                            p++; if (p < source.Length && (source[p] == '+' || source[p] == '-')) p++;
                            int digits = p; while (p < source.Length && Char.IsDigit(source[p])) p++;
                            if (digits == p) throw Error("Invalid numeric exponent", start);
                        }
                        // Sheet names may begin with a digit, e.g. 2026!A1.
                        if (p < source.Length && (source[p] == '!' || (IsNameCharacter(source[p]) && !IsStyledSymbol(source[p]))))
                        {
                            p = start; ScanAtom(ref p); kind = "Atom";
                            if (source.Substring(start, p - start).IndexOf('!') < 0) throw Error("Invalid number or identifier beginning with a digit", start);
                        }
                    }
                    else if (c == '#')
                    {
                        if (TryScanError(ref p)) kind = "Error";
                        else p++;
                    }
                    else if (IsStyledSymbol(c))
                    {
                        if (c == displayStyle.ArrayOpen) arrayDepth++;
                        else if (c == displayStyle.ArrayClose) arrayDepth--;
                        p++;
                    }
                    else if (c == '\'' || c == '[' || IsNameCharacter(c)) { ScanAtom(ref p); kind = "Atom"; }
                    else if ("+-*/^&=<>%:,;(){}@".IndexOf(c) >= 0)
                    {
                        p++;
                        if (p < source.Length && ((c == '<' && (source[p] == '>' || source[p] == '=')) || (c == '>' && source[p] == '='))) p++;
                    }
                    else throw Error(c == '|' ? "DDE expressions are unsupported" : "Unsupported character '" + c + "'", p);
                    tokens.Add(new Token { Start = start, End = p, Kind = kind, Text = source.Substring(start, p - start) });
                }
                tokens.Add(new Token { Start = source.Length, End = source.Length, Kind = "End", Text = "" });
            }
            static bool IsNameCharacter(char c) { return Char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '$' || c == '\\' || c == '?'; }

            bool IsStyledSymbol(char c)
            {
                return displayStyle != null && (c == displayStyle.ArrayOpen || c == displayStyle.ArrayClose ||
                    c == displayStyle.ListSeparator || (arrayDepth > 0 && (c == displayStyle.ArrayColumnSeparator || c == displayStyle.ArrayRowSeparator)));
            }

            bool TryScanError(ref int p)
            {
                foreach (string error in Errors)
                    if (source.Length - p >= error.Length && String.Compare(source, p, error, 0, error.Length, StringComparison.OrdinalIgnoreCase) == 0)
                    { p += error.Length; return true; }
                if (displayStyle == null || !displayStyle.Local) return false;
                int end = p + 1;
                if (end >= source.Length || !Char.IsLetter(source[end])) return false;
                while (end < source.Length && (Char.IsLetterOrDigit(source[end]) || "_./".IndexOf(source[end]) >= 0)) end++;
                if (end < source.Length && (source[end] == '!' || source[end] == '?')) end++;
                p = end;
                return true;
            }

            bool TryScanStyledGrid(ref int p)
            {
                if (displayGridReference == null) return false;
                Match match = displayGridReference.Match(source, p);
                if (!match.Success || match.Index != p) return false;
                int end = p + match.Length;
                if (end < source.Length && (IsNameCharacter(source[end]) || source[end] == '!' || source[end] == '[')) return false;
                p = end;
                return true;
            }

            void ScanAtom(ref int p)
            {
                int start = p;
                bool qualified = false;
                if (TryScanStyledGrid(ref p)) return;
                if (source[p] == '\'')
                {
                    p++; bool closed = false;
                    while (p < source.Length)
                    {
                        if (source[p++] != '\'') continue;
                        if (p < source.Length && source[p] == '\'') { p++; continue; }
                        closed = true; break;
                    }
                    if (!closed) throw Error("Unclosed quoted sheet or workbook name", start);
                    if (p >= source.Length || source[p] != '!') throw Error("Quoted sheet or workbook name must be followed by '!'", p);
                    p++; qualified = true;
                    if (TryScanStyledGrid(ref p)) return;
                }
                while (p < source.Length)
                {
                    if (IsStyledSymbol(source[p])) break;
                    if (IsNameCharacter(source[p])) { p++; continue; }
                    if (source[p] == '!')
                    {
                        if (qualified) throw Error("A reference has more than one sheet separator", p);
                        qualified = true; p++;
                        if (TryScanStyledGrid(ref p)) return;
                        continue;
                    }
                    if (qualified && source[p] == '#' && displayStyle != null && TryScanError(ref p)) return;
                    if (qualified && source.Length - p >= 5 && String.Compare(source, p, "#REF!", 0, 5, StringComparison.OrdinalIgnoreCase) == 0) { p += 5; break; }
                    if (source[p] != '[') break;
                    int bracketStart = p, brackets = 0;
                    do
                    {
                        char current = source[p++];
                        if (current == '\'' && p < source.Length && "[]#'@".IndexOf(source[p]) >= 0) { p++; continue; }
                        if (current == '[') brackets++;
                        if (current == ']') brackets--;
                    } while (p < source.Length && brackets != 0);
                    if (brackets != 0) throw Error("Unclosed structured or external reference", bracketStart);
                }
                if (p == start || (source[p - 1] == '!' && !source.Substring(start, p - start).EndsWith("#REF!", StringComparison.OrdinalIgnoreCase))) throw Error("Incomplete reference", start);
            }
        }
    }
}
