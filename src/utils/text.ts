import { createHash } from 'node:crypto';

/** Strips HTML to readable plain text without pulling in a DOM dependency. */
export function htmlToText(html: string | null | undefined): string {
  if (!html) return '';
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n')
    .replace(/<li[^>]*>/gi, '- ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
/** Drops Markdown emphasis markers that senders paste into mail bodies. */
export function stripEmphasis(text: string): string {
  return text.replace(/\*\*(.+?)\*\*/gs, '$1').replace(/(^|\s)\*(\S[^*]*?)\*(?=\s|$)/g, '$1$2');
}

const QUOTE_MARKERS = [
  /^\s*-{2,}\s*original message\s*-{2,}/im,
  /^\s*_{5,}\s*$/m,
  /^\s*from:\s.+\bsent:\s/im,
  /^\s*差出人:\s/im,
  /^\s*送信元:\s/im,
  /^\s*>\s?.+$/m,
  /^\s*On .+ wrote:\s*$/im,
];

/** Removes quoted history so triage only sees what the sender actually wrote. */
export function stripQuotedHistory(text: string): string {
  let cut = text.length;
  for (const marker of QUOTE_MARKERS) {
    const match = marker.exec(text);
    if (match?.index !== undefined && match.index > 0 && match.index < cut) {
      cut = match.index;
    }
  }
  const trimmed = text.slice(0, cut).trim();
  return trimmed.length >= 8 ? trimmed : text.trim();
}

/** Normalises a mail subject by removing Re:/Fwd:/RE: prefixes and noise. */
export function normalizeSubject(subject: string | null | undefined): string {
  if (!subject) return '';
  let value = subject;
  let changed = true;
  while (changed) {
    const next = value
      .replace(/^\s*(re|fw|fwd|返信|転送|回答|自動転送)\s*(\[\d+\])?\s*[:：]\s*/i, '')
      .replace(/^\s*\[(external|外部)\]\s*/i, '');
    changed = next !== value;
    value = next;
  }
  return value.replace(/\s+/g, ' ').trim();
}

const JAPANESE = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uff66-\uff9f]/;

export type Language = 'ja' | 'en';

/** Detects the dominant language so replies are drafted in kind (spec 15). */
export function detectLanguage(text: string): Language {
  if (!text) return 'ja';
  const japaneseChars = (text.match(/[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uff66-\uff9f]/g) ?? []).length;
  const latinChars = (text.match(/[A-Za-z]/g) ?? []).length;
  if (japaneseChars === 0) return 'en';
  // A handful of Japanese characters in an otherwise English mail (signatures,
  // names) should not flip the reply language.
  return japaneseChars * 6 >= latinChars ? 'ja' : 'en';
}

export function containsJapanese(text: string): boolean {
  return JAPANESE.test(text);
}

export function truncate(text: string, max: number): string {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (clean.length <= max) return clean;
  return `${clean.slice(0, Math.max(0, max - 1))}…`;
}

export function shortHash(value: string): string {
  return createHash('sha256').update(value).digest('hex').slice(0, 16);
}

/** Loose token overlap used by the de-duplication pass (spec 16). */
export function similarity(a: string, b: string): number {
  const tokenise = (value: string) =>
    new Set(
      value
        .toLowerCase()
        .replace(/[^\p{L}\p{N}\s]/gu, ' ')
        .split(/\s+/)
        .filter((token) => token.length > 1),
    );
  const setA = tokenise(a);
  const setB = tokenise(b);
  if (setA.size === 0 || setB.size === 0) {
    // Fall back to character bigrams for Japanese text, which does not split on
    // whitespace.
    return bigramSimilarity(a, b);
  }
  let intersection = 0;
  for (const token of setA) if (setB.has(token)) intersection += 1;
  const union = setA.size + setB.size - intersection;
  const jaccard = union === 0 ? 0 : intersection / union;
  return Math.max(jaccard, bigramSimilarity(a, b));
}

function bigramSimilarity(a: string, b: string): number {
  const grams = (value: string) => {
    const clean = value.replace(/\s+/g, '');
    const result = new Set<string>();
    for (let i = 0; i < clean.length - 1; i += 1) result.add(clean.slice(i, i + 2));
    return result;
  };
  const gramsA = grams(a);
  const gramsB = grams(b);
  if (gramsA.size === 0 || gramsB.size === 0) return 0;
  let intersection = 0;
  for (const gram of gramsA) if (gramsB.has(gram)) intersection += 1;
  return (2 * intersection) / (gramsA.size + gramsB.size);
}
