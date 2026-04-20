"use strict";

// UTC 基準の日付ユーティリティ。
// Power BI は日時値を UTC epoch の Date / ISO 文字列で渡すため、
// getUTCFullYear/Month/Date で日付部分を取り出してローカル TZ ズレを回避する。

/** Date → "YYYY-MM-DD"（UTC） */
export function formatDateUTC(d: Date): string {
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
}

/** "YYYY-MM-DD" → UTC 0:00 epoch。不正値は NaN */
export function toDateEpochFromString(s: string): number {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return NaN;
    const [y, m, d] = s.split("-").map(Number);
    return Date.UTC(y, m - 1, d);
}

/**
 * 任意形式の値 → UTC 0:00 epoch（時刻切り捨て）。
 * Date / ISO 文字列 / 非 ISO 文字列 / 数値 epoch ms に対応。
 */
export function toDateEpoch(v: unknown): number {
    if (v == null) return NaN;
    if (v instanceof Date) {
        const t = v.getTime();
        if (!Number.isFinite(t)) return NaN;
        return Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
    }
    if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
        const d = new Date(v);
        const t = d.getTime();
        if (Number.isFinite(t)) return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
        return NaN;
    }
    if (typeof v === "number" && Number.isFinite(v)) {
        const d = new Date(v);
        return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    return NaN;
}

/** epoch → 翌日 0:00 の epoch */
export function nextDayEpoch(epoch: number): number {
    return epoch + 86400000;
}

/** "YYYY-MM-DD" が有効な日付か */
export function isValidYMD(s: string): boolean {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
    const ep = toDateEpochFromString(s);
    if (!Number.isFinite(ep)) return false;
    // 月末日チェック（例: 2026-02-30 を弾く）
    const d = new Date(ep);
    return formatDateUTC(d) === s;
}

export const MONTH_NAMES_JA = [
    "1月", "2月", "3月", "4月", "5月", "6月",
    "7月", "8月", "9月", "10月", "11月", "12月",
];

export const WEEKDAY_NAMES_JA = ["日", "月", "火", "水", "木", "金", "土"];
