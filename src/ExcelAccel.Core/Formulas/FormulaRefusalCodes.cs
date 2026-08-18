namespace ExcelAccel.Core.Formulas;

public static class FormulaRefusalCodes
{
    public const string Empty = "FORMULA_EMPTY";
    public const string PrefixRequired = "FORMULA_PREFIX_REQUIRED";
    public const string TooLong = "FORMULA_TOO_LONG";
    public const string TooManyTokens = "FORMULA_TOO_MANY_TOKENS";
    public const string NestingLimit = "FORMULA_NESTING_LIMIT";
    public const string UnbalancedDelimiter = "FORMULA_UNBALANCED_DELIMITER";
    public const string UnterminatedString = "FORMULA_UNTERMINATED_STRING";
    public const string UnsupportedCharacter = "FORMULA_UNSUPPORTED_CHARACTER";
    public const string UnsupportedArraySyntax = "FORMULA_ARRAY_SYNTAX_UNSUPPORTED";
    public const string DialectMismatch = "FORMULA_DIALECT_MISMATCH";
    public const string InvalidReference = "FORMULA_INVALID_REFERENCE";
    public const string InvalidSyntax = "FORMULA_INVALID_SYNTAX";
    public const string StructuredReferenceInspectOnly = "FORMULA_STRUCTURED_REFERENCE_INSPECT_ONLY";
    public const string ExternalReferenceInspectOnly = "FORMULA_EXTERNAL_REFERENCE_INSPECT_ONLY";
    public const string NameInspectOnly = "FORMULA_NAME_INSPECT_ONLY";
    public const string DynamicArrayInspectOnly = "FORMULA_DYNAMIC_ARRAY_INSPECT_ONLY";
    public const string IntersectionInspectOnly = "FORMULA_INTERSECTION_INSPECT_ONLY";
    public const string UnionInspectOnly = "FORMULA_UNION_INSPECT_ONLY";
}
