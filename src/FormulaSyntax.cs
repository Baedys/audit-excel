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
        public override string ToString() { return Text; }
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
        {
            if (root == null || target == null) throw new ArgumentNullException(root == null ? "root" : "target");
            var path = new List<ExpressionNode>();
            if (!FindPath(root, target, path)) throw new ArgumentException("The selected expression does not belong to this formula.", "target");
            string expression = target.Text;
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

        static string FunctionName(ExpressionNode node)
        {
            if (node.Kind != "Function") return "";
            int parenthesis = node.Text.IndexOf('(');
            string name = parenthesis < 0 ? "" : node.Text.Substring(0, parenthesis).Trim().ToUpperInvariant();
            return name.StartsWith("_XLFN.", StringComparison.Ordinal) ? name.Substring(6) : name;
        }

        static bool FindPath(ExpressionNode node, ExpressionNode target, List<ExpressionNode> path)
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

        sealed class Token
        {
            public string Text, Kind;
            public int Start, End;
        }

        sealed class Parser
        {
            readonly string source;
            readonly List<Token> tokens = new List<Token>();
            int index, depth;
            static readonly Regex R1C1Token = new Regex(@"^R(?:\d+|\[-?\d+\])?C(?:\d+|\[-?\d+\])?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            static readonly string[] Errors = { "#GETTING_DATA", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#N/A", "#SPILL!", "#CALC!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#BUSY!", "#PYTHON!" };

            public Parser(string text) { source = text; Lex(); }
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
                return node.Kind == "Reference" || node.Kind == "Name" || node.Kind == "Function" || node.Kind == "Invocation" || node.Kind == "Range" || node.Kind == "Union" || node.Kind == "Intersection" || node.Kind == "Spill";
            }
            static bool RangeEndpoint(ExpressionNode node)
            {
                return ReferenceLike(node) || (node.Kind == "Number" && Regex.IsMatch(node.Text, @"^[1-9][0-9]*$"));
            }
            static bool StartsReference(Token token) { return token.Kind == "Atom" || token.Text == "("; }

            ExpressionNode Prefix(bool union)
            {
                Token token = Take();
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
                while (p < source.Length)
                {
                    if (Char.IsWhiteSpace(source[p])) { p++; continue; }
                    int start = p;
                    char c = source[p];
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
                    else if (Char.IsDigit(c) || (c == '.' && p + 1 < source.Length && Char.IsDigit(source[p + 1])))
                    {
                        kind = "Number";
                        while (p < source.Length && Char.IsDigit(source[p])) p++;
                        if (p < source.Length && source[p] == '.') { p++; while (p < source.Length && Char.IsDigit(source[p])) p++; }
                        if (p + 1 < source.Length && (source[p] == 'E' || source[p] == 'e') && (Char.IsDigit(source[p + 1]) || source[p + 1] == '+' || source[p + 1] == '-'))
                        {
                            p++; if (p < source.Length && (source[p] == '+' || source[p] == '-')) p++;
                            int digits = p; while (p < source.Length && Char.IsDigit(source[p])) p++;
                            if (digits == p) throw Error("Invalid numeric exponent", start);
                        }
                        // Sheet names may begin with a digit, e.g. 2026!A1.
                        if (p < source.Length && (source[p] == '!' || IsNameCharacter(source[p])))
                        {
                            p = start; ScanAtom(ref p); kind = "Atom";
                            if (source.Substring(start, p - start).IndexOf('!') < 0) throw Error("Invalid number or identifier beginning with a digit", start);
                        }
                    }
                    else if (c == '#')
                    {
                        bool found = false;
                        foreach (string error in Errors)
                        {
                            if (source.Length - p >= error.Length && String.Compare(source, p, error, 0, error.Length, StringComparison.OrdinalIgnoreCase) == 0)
                            { p += error.Length; kind = "Error"; found = true; break; }
                        }
                        if (!found) p++;
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

            void ScanAtom(ref int p)
            {
                int start = p;
                bool qualified = false;
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
                }
                while (p < source.Length)
                {
                    if (IsNameCharacter(source[p])) { p++; continue; }
                    if (source[p] == '!')
                    {
                        if (qualified) throw Error("A reference has more than one sheet separator", p);
                        qualified = true; p++; continue;
                    }
                    if (qualified && source.Length - p >= 5 && String.Compare(source, p, "#REF!", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
                    { p += 5; break; }
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
