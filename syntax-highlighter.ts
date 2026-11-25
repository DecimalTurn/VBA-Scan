export type SourceLanguage = "vba" | "pcode";

export interface SyntaxToken {
  text: string;
  kind: string | null;
}

export interface SyntaxHighlightBatchMetrics {
  sourceChars: number;
  tokenCount: number;
  highlightedTokenCount: number;
  tokenizeMs: number;
  renderMs: number;
  totalMs: number;
}

const keywordList = [
  "addressOf", "alias", "and", "as", "base", "begin", "byref", "byval",
  "call", "case", "class", "const", "declare", "dim", "do", "doevents",
  "each", "else", "elseif", "end", "enum", "eqv", "erase", "exit",
  "explicit", "for", "friend", "function", "get", "global", "gosub",
  "goto", "if", "imp", "implements", "in", "input", "is", "let", "lib",
  "like", "lock", "loop", "mod", "new", "next", "not", "on", "open",
  "option", "optional", "or", "output", "paramarray", "preserve", "print",
  "private", "property", "ptrsafe", "public", "raiseevent", "random",
  "redim", "resume", "return", "select", "set", "shared", "static",
  "step", "stop", "sub", "then", "to", "typeof", "until", "wend", "while",
  "with", "withevents", "write", "xor",
];

const typeList = [
  "boolean", "byte", "currency", "date", "decimal", "double", "integer",
  "long", "longlong", "longptr", "object", "single", "string", "variant",
];

const constantList = [
  "false", "nothing", "true", "vbcr", "vbcrlf", "vblf", "vbnewline",
  "vba7", "win64",
];

function alternation(values: string[]): string {
  return values.map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
}

function compileTokenizer(source: string, flags = "gim"): RegExp {
  return new RegExp(source, flags);
}

/*
 * VBA lexical categories are adapted from Sam Vanderslink's VBA-LanguageServer
 * TextMate grammar (MIT). See assets/VBA-LanguageServer.LICENSE.txt.
 */
const vbaTokenizer = compileTokenizer([
  String.raw`(?<string>"(?:""|[^"\r\n])*")`,
  String.raw`(?<comment>'[^\r\n]*|\brem\b[^\r\n]*)`,
  String.raw`(?<directive>#(?:if|elseif|else|end\s+if|const|then)\b)`,
  String.raw`(?<attribute>^\s*(?:attribute|version)\b)`,
  String.raw`(?<hexadecimal>&[hH][0-9a-f]+&?)`,
  String.raw`(?<number>(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?[!#%&@]?)`,
  String.raw`(?<type>\b(?:${alternation(typeList)})\b)`,
  String.raw`(?<constant>\b(?:${alternation(constantList)})\b)`,
  String.raw`(?<keyword>\b(?:${alternation(keywordList)})\b)`,
  String.raw`(?<function>[a-z_][a-z0-9_$]*(?=\s*\())`,
  String.raw`(?<operator>:=|<>|<=|>=|\b(?:addressof|typeof)\b|[&*+\-./^<=>:])`,
  String.raw`(?<identifier>[a-z_][a-z0-9_$]*)`,
  String.raw`(?<space>[\t ]+|\r?\n)`,
  String.raw`(?<plain>.)`,
].join("|"));

const pcodeTokenizer = compileTokenizer([
  String.raw`(?<pcodeHeader>^(?:={3,}[^\r\n]*|module streams:)[^\r\n]*$)`,
  String.raw`(?<pcodeModule>^VBA\/[^\r\n]+ - \d+ bytes$)`,
  String.raw`(?<pcodeLine>^line #\d+:$)`,
  String.raw`(?<string>"(?:""|[^"\r\n])*")`,
  String.raw`(?<hexadecimal>\b0x[0-9a-f]+\b)`,
  String.raw`(?<number>\b\d+(?:\.\d+)?\b)`,
  String.raw`(?<instruction>^[\t ]+[A-Z][A-Za-z0-9_$]*(?=[\t \r\n]|$))`,
  String.raw`(?<reference><[^>\r\n]+>)`,
  String.raw`(?<identifier>[A-Za-z_][A-Za-z0-9_$]*)`,
  String.raw`(?<space>[\t ]+|\r?\n)`,
  String.raw`(?<plain>.)`,
].join("|"));

const tokenizers: Record<SourceLanguage, RegExp> = {
  vba: vbaTokenizer,
  pcode: pcodeTokenizer,
};

export function tokenizeCode(source: string, language: SourceLanguage): SyntaxToken[] {
  const tokenizer = tokenizers[language];
  const tokens: SyntaxToken[] = [];
  tokenizer.lastIndex = 0;
  for (const match of source.matchAll(tokenizer)) {
    const groups = match.groups ?? {};
    const entry = Object.entries(groups).find(([, value]) => value !== undefined);
    const kind = entry?.[0].replace(/([a-z])([A-Z])/g, "$1-$2").toLowerCase() ?? null;
    tokens.push({
      text: match[0],
      kind: kind === "plain" || kind === "space" ? null : kind,
    });
  }
  return tokens;
}

export function renderHighlightedCode(
  container: HTMLElement,
  source: string,
  language: SourceLanguage,
): void {
  const fragment = document.createDocumentFragment();
  for (const token of tokenizeCode(source, language)) {
    if (!token.kind) {
      fragment.append(document.createTextNode(token.text));
      continue;
    }
    const span = document.createElement("span");
    span.className = `syntax-token syntax-${token.kind}`;
    span.textContent = token.text;
    fragment.append(span);
  }
  container.replaceChildren(fragment);
}

export class StreamingSyntaxHighlighter {
  private pending = "";
  private pendingNode: Text | null = null;

  constructor(
    private readonly container: HTMLElement,
    private readonly language: SourceLanguage,
  ) {}

  reset(): void {
    this.pending = "";
    this.pendingNode = null;
  }

  append(chunk: string, final = false): SyntaxHighlightBatchMetrics {
    const startedAt = performance.now();
    this.pendingNode?.remove();
    this.pendingNode = null;
    const text = this.pending + chunk;
    const lastNewline = final ? text.length : text.lastIndexOf("\n") + 1;
    const complete = lastNewline > 0 ? text.slice(0, lastNewline) : "";
    this.pending = lastNewline > 0 ? text.slice(lastNewline) : text;
    const tokenizeStartedAt = performance.now();
    const tokens = complete ? tokenizeCode(complete, this.language) : [];
    const tokenizeMs = performance.now() - tokenizeStartedAt;
    const renderStartedAt = performance.now();
    let highlightedTokenCount = 0;
    if (complete) {
      const fragment = document.createDocumentFragment();
      for (const token of tokens) {
        if (!token.kind) {
          fragment.append(document.createTextNode(token.text));
          continue;
        }
        highlightedTokenCount++;
        const span = document.createElement("span");
        span.className = `syntax-token syntax-${token.kind}`;
        span.textContent = token.text;
        fragment.append(span);
      }
      this.container.append(fragment);
    }
    if (this.pending) {
      this.pendingNode = document.createTextNode(this.pending);
      this.container.append(this.pendingNode);
    }
    const renderMs = performance.now() - renderStartedAt;
    return {
      sourceChars: complete.length,
      tokenCount: tokens.length,
      highlightedTokenCount,
      tokenizeMs,
      renderMs,
      totalMs: performance.now() - startedAt,
    };
  }
}
