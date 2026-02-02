// Avanti PPT Anonymizer — Configuration & constants

export const SUPABASE_URL = 'https://vnjcwffdhywckwnjothu.supabase.co';
export const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZuamN3ZmZkaHl3Y2t3bmpvdGh1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU4NjA4MTAsImV4cCI6MjA4MTQzNjgxMH0.ETCptr-BYt7wunTOXVAsBCsv9L9kICR30GGHoC5X3ZQ';

export const SECTION_LABELS = new Set([
    'syfte', 'mål', 'bakgrund', 'agenda', 'sammanfattning',
    'nästa steg', 'tidplan', 'organisation', 'risker',
    'budget', 'resurser', 'bilagor', 'innehåll', 'analys',
    'resultat', 'slutsats', 'rekommendationer', 'översikt',
    'introduktion', 'diskussion', 'metod', 'uppföljning'
]);

export const PLACEHOLDER_MAP = {
    title:       '[Rubrik]',
    body:        '[Beskrivning]',
    name:        '[Namn]',
    number:      '[Värde]',
    initials:    '[Initialer]',
    email:       '[email]',
    phone:       '[telefon]',
    url:         '[URL]',
    date:        '[Period]',
};

export const LOCAL_PATTERNS = [
    {
        category: 'email',
        test: (text) => /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(text.trim())
    },
    {
        category: 'phone',
        test: (text) => /^(?:\+46|0)[\s-]?(?:\d[\s-]?){8,11}$/.test(text.trim())
    },
    {
        category: 'url',
        test: (text) => /^https?:\/\/[^\s]+$/.test(text.trim())
    },
    {
        category: 'initials',
        test: (text) => /^[A-ZÅÄÖ]{2,3}(?:\s*[+&,]\s*[A-ZÅÄÖ]{2,3})+$/.test(text.trim())
    },
    {
        category: 'date',
        test: (text) => /^(?:Q[1-4]\s*\d{4}|\d{4}-\d{2}-\d{2}|\d{4}-\d{2}|\d{1,2}\/\d{1,2}[\/-]\d{2,4}|(?:jan|feb|mar|apr|maj|jun|jul|aug|sep|okt|nov|dec)\w*\s+\d{4})$/i.test(text.trim())
    },
    {
        category: 'number',
        test: (text) => /^\d+(?:[,.\s]\d{3})*(?:[,.\s]\d+)?\s*(?:kr|SEK|MSEK|TSEK|EUR|USD|miljoner|miljarder|st|%|kkr|mkr|mdr)$/i.test(text.trim())
    },
    {
        category: 'section_label',
        test: (text) => SECTION_LABELS.has(text.trim().toLowerCase())
    }
];

export function makeKey(slideIndex, shapeIndex, row, col, groupChildIndex) {
    let key = `${slideIndex}-${shapeIndex}`;
    if (groupChildIndex !== undefined) key += `-g${groupChildIndex}`;
    if (row !== undefined) key += `-${row}-${col}`;
    return key;
}

export function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
