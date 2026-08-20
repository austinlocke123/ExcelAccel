using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Formulas;

public enum FormulaNodeKind
{
    Function,
    BinaryOperator,
    UnaryOperator,
    PostfixOperator,
    Group,
    Reference,
    Name,
    Number,
    Text,
    Boolean,
    ErrorLiteral,
}

/// <summary>
/// One immutable node of a parsed formula. A node knows what it is, the exact
/// source span it came from, and its children. It carries no value: the tree is
/// never evaluated.
/// </summary>
public sealed class FormulaSyntaxNode
{
    internal FormulaSyntaxNode(FormulaNodeKind kind, string text, FormulaSourceSpan span, IReadOnlyList<FormulaSyntaxNode> children)
    {
        Kind = kind;
        Text = text;
        Span = span;
        Children = children;
    }

    public FormulaNodeKind Kind { get; }

    /// <summary>The function name, operator symbol, or literal text.</summary>
    public string Text { get; }

    /// <summary>The exact span of source this node covers.</summary>
    public FormulaSourceSpan Span { get; }

    public IReadOnlyList<FormulaSyntaxNode> Children { get; }

    public bool IsLeaf => Children.Count == 0;

    /// <summary>Pre-order flattening, which is also the keyboard focus order.</summary>
    public IEnumerable<(FormulaSyntaxNode Node, int Depth)> Flatten(int depth = 0)
    {
        yield return (this, depth);
        foreach (var child in Children)
        {
            foreach (var descendant in child.Flatten(depth + 1)) yield return descendant;
        }
    }
}

public sealed class FormulaTreeResult
{
    private FormulaTreeResult(FormulaSyntaxNode? root, int nodeCount, string? limitationCode, string? message, FormulaSourceSpan? limitationSpan)
    {
        Root = root;
        NodeCount = nodeCount;
        LimitationCode = limitationCode;
        Message = message;
        LimitationSpan = limitationSpan;
    }

    public FormulaSyntaxNode? Root { get; }

    public int NodeCount { get; }

    /// <summary>Set when the formula could not be fully represented as a tree.</summary>
    public string? LimitationCode { get; }

    public string? Message { get; }

    /// <summary>The exact span the limitation applies to, when one is known.</summary>
    public FormulaSourceSpan? LimitationSpan { get; }

    public bool IsComplete => Root is not null && LimitationCode is null;

    internal static FormulaTreeResult Success(FormulaSyntaxNode root, int nodeCount) =>
        new FormulaTreeResult(root, nodeCount, null, null, null);

    internal static FormulaTreeResult Limited(string code, string message, FormulaSourceSpan? span, FormulaSyntaxNode? partialRoot = null, int nodeCount = 0) =>
        new FormulaTreeResult(partialRoot, nodeCount, code, message, span);

    /// <summary>
    /// A formula the qualified parser refused outright. The inspector reports the
    /// parser's own reason rather than inventing one.
    /// </summary>
    public static FormulaTreeResult Refused(string? refusalCode, string? message) =>
        new FormulaTreeResult(
            null,
            0,
            string.IsNullOrWhiteSpace(refusalCode) ? FormulaRefusalCodes.InvalidSyntax : refusalCode!,
            string.IsNullOrWhiteSpace(message) ? "The formula is outside qualified parser coverage." : message!,
            null);
}

/// <summary>
/// Builds an immutable syntax tree over the qualified token stream.
///
/// This is additive: the parser's own tokens, references, coverage disposition,
/// and limitation codes are unchanged, and a formula the parser already refuses
/// never reaches the builder. A construct the builder cannot represent yields an
/// explicit limitation with its exact span rather than a tree that looks
/// complete.
///
/// The tree is structural only. It is never evaluated, scored, or explained.
/// </summary>
public static class FormulaTreeBuilder
{
    /// <summary>Node ceiling, so a pathological formula cannot exhaust the view.</summary>
    public const int MaximumNodes = 2_000;

    public static FormulaTreeResult Build(FormulaSyntaxDocument document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        // A construct the qualified parser marks inspect-only for a structural
        // reason cannot be represented faithfully as a tree.
        foreach (var code in document.LimitationCodes)
        {
            if (code == FormulaRefusalCodes.StructuredReferenceInspectOnly ||
                code == FormulaRefusalCodes.DynamicArrayInspectOnly ||
                code == FormulaRefusalCodes.IntersectionInspectOnly ||
                code == FormulaRefusalCodes.UnionInspectOnly)
            {
                return FormulaTreeResult.Limited(
                    code,
                    "This formula uses a construct the qualified parser does not represent as a tree.",
                    new FormulaSourceSpan(0, document.SourceText.Length));
            }
        }

        var tokens = document.Tokens.Where(token => token.Kind != FormulaTokenKind.Whitespace).ToArray();
        var start = tokens.Length != 0 && tokens[0].Kind == FormulaTokenKind.Prefix ? 1 : 0;
        if (start >= tokens.Length)
        {
            return FormulaTreeResult.Limited(FormulaRefusalCodes.Empty, "The formula has no expression to inspect.", null);
        }

        var state = new ParseState(tokens, start);
        FormulaSyntaxNode root;
        try
        {
            root = ParseExpression(state, 0);
        }
        catch (TreeLimitationException exception)
        {
            return FormulaTreeResult.Limited(exception.Code, exception.Message, exception.Span);
        }

        if (state.Index < tokens.Length)
        {
            return FormulaTreeResult.Limited(
                FormulaRefusalCodes.InvalidSyntax,
                "The formula has trailing content the inspector could not attach to the tree.",
                tokens[state.Index].Span);
        }

        var count = root.Flatten().Count();
        if (count > MaximumNodes)
        {
            return FormulaTreeResult.Limited(
                FormulaRefusalCodes.TooManyTokens,
                $"The formula expands to more than {MaximumNodes:N0} nodes.",
                new FormulaSourceSpan(0, document.SourceText.Length),
                root,
                count);
        }

        return FormulaTreeResult.Success(root, count);
    }

    /// <summary>Excel operator precedence, lowest binding first.</summary>
    private static int Precedence(string op) => op switch
    {
        "=" or "<>" or "<" or ">" or "<=" or ">=" => 1,
        "&" => 2,
        "+" or "-" => 3,
        "*" or "/" => 4,
        "^" => 5,
        _ => -1,
    };

    private static FormulaSyntaxNode ParseExpression(ParseState state, int minimumPrecedence)
    {
        var left = ParseUnary(state);
        while (state.Current is { Kind: FormulaTokenKind.Operator } token)
        {
            var precedence = Precedence(token.Text);
            if (precedence < 0 || precedence < minimumPrecedence) break;
            state.Advance();
            // "^" is right-associative in Excel; everything else is left.
            var right = ParseExpression(state, token.Text == "^" ? precedence : precedence + 1);
            left = new FormulaSyntaxNode(FormulaNodeKind.BinaryOperator, token.Text, Cover(left.Span, right.Span), new[] { left, right });
        }

        return left;
    }

    private static FormulaSyntaxNode ParseUnary(ParseState state)
    {
        if (state.Current is { Kind: FormulaTokenKind.Operator } token && (token.Text == "-" || token.Text == "+"))
        {
            state.Advance();
            var operand = ParseUnary(state);
            return new FormulaSyntaxNode(FormulaNodeKind.UnaryOperator, token.Text, Cover(token.Span, operand.Span), new[] { operand });
        }

        return ParsePostfix(state);
    }

    private static FormulaSyntaxNode ParsePostfix(ParseState state)
    {
        var node = ParsePrimary(state);
        while (state.Current is { Kind: FormulaTokenKind.Operator } token && token.Text == "%")
        {
            state.Advance();
            node = new FormulaSyntaxNode(FormulaNodeKind.PostfixOperator, "%", Cover(node.Span, token.Span), new[] { node });
        }

        return node;
    }

    private static FormulaSyntaxNode ParsePrimary(ParseState state)
    {
        var token = state.Current ?? throw new TreeLimitationException(
            FormulaRefusalCodes.InvalidSyntax, "The formula ends before its expression is complete.", null);

        switch (token.Kind)
        {
            case FormulaTokenKind.Number:
                state.Advance();
                return Leaf(FormulaNodeKind.Number, token);
            case FormulaTokenKind.StringLiteral:
                state.Advance();
                return Leaf(FormulaNodeKind.Text, token);
            case FormulaTokenKind.ErrorLiteral:
                state.Advance();
                return Leaf(FormulaNodeKind.ErrorLiteral, token);
            case FormulaTokenKind.Reference:
            case FormulaTokenKind.QuotedIdentifier:
                state.Advance();
                return Leaf(FormulaNodeKind.Reference, token);
            case FormulaTokenKind.OpenParenthesis:
                {
                    state.Advance();
                    var inner = ParseExpression(state, 0);
                    var close = Expect(state, FormulaTokenKind.CloseParenthesis, "A parenthesised group is not closed.");
                    return new FormulaSyntaxNode(FormulaNodeKind.Group, "()", Cover(token.Span, close.Span), new[] { inner });
                }
            case FormulaTokenKind.Identifier:
                {
                    if (IsBoolean(token.Text))
                    {
                        state.Advance();
                        return Leaf(FormulaNodeKind.Boolean, token);
                    }

                    state.Advance();
                    if (state.Current is { Kind: FormulaTokenKind.OpenParenthesis })
                    {
                        state.Advance();
                        var arguments = new List<FormulaSyntaxNode>();
                        if (state.Current is not { Kind: FormulaTokenKind.CloseParenthesis })
                        {
                            arguments.Add(ParseExpression(state, 0));
                            while (state.Current is { Kind: FormulaTokenKind.Separator })
                            {
                                state.Advance();
                                arguments.Add(ParseExpression(state, 0));
                            }
                        }

                        var close = Expect(state, FormulaTokenKind.CloseParenthesis, "A function call is not closed.");
                        return new FormulaSyntaxNode(FormulaNodeKind.Function, token.Text, Cover(token.Span, close.Span), arguments.AsReadOnly());
                    }

                    return Leaf(FormulaNodeKind.Name, token);
                }
            default:
                throw new TreeLimitationException(
                    FormulaRefusalCodes.InvalidSyntax,
                    $"The inspector does not represent a {token.Kind} token as a tree node.",
                    token.Span);
        }
    }

    private static FormulaToken Expect(ParseState state, FormulaTokenKind kind, string message)
    {
        var token = state.Current;
        if (token is null || token.Kind != kind)
        {
            throw new TreeLimitationException(FormulaRefusalCodes.UnbalancedDelimiter, message, token?.Span);
        }

        state.Advance();
        return token;
    }

    private static FormulaSyntaxNode Leaf(FormulaNodeKind kind, FormulaToken token) =>
        new FormulaSyntaxNode(kind, token.Text, token.Span, Array.Empty<FormulaSyntaxNode>());

    private static FormulaSourceSpan Cover(FormulaSourceSpan first, FormulaSourceSpan second)
    {
        var start = Math.Min(first.Start, second.Start);
        var end = Math.Max(first.Start + first.Length, second.Start + second.Length);
        return new FormulaSourceSpan(start, end - start);
    }

    private static bool IsBoolean(string value) =>
        string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase);

    private sealed class ParseState
    {
        private readonly IReadOnlyList<FormulaToken> _tokens;

        public ParseState(IReadOnlyList<FormulaToken> tokens, int index)
        {
            _tokens = tokens;
            Index = index;
        }

        public int Index { get; private set; }

        public FormulaToken? Current => Index < _tokens.Count ? _tokens[Index] : null;

        public void Advance() => Index++;
    }

    private sealed class TreeLimitationException : Exception
    {
        public TreeLimitationException(string code, string message, FormulaSourceSpan? span)
            : base(message)
        {
            Code = code;
            Span = span;
        }

        public string Code { get; }

        public FormulaSourceSpan? Span { get; }
    }
}

