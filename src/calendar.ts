"use strict";

import {
    formatDateUTC,
    toDateEpochFromString,
    MONTH_NAMES_JA,
    WEEKDAY_NAMES_JA,
} from "./dateUtils";

export type CalendarMode = "single" | "range";

export interface CalendarState {
    mode: CalendarMode;
    /** "YYYY-MM-DD" 形式。単一モードでは startDate のみ使用 */
    startDate: string;
    endDate: string;
}

export interface CalendarCallbacks {
    /** 選択が確定したとき（単一: 1 クリック / 範囲: 2 クリック目）に発火 */
    onChange: (state: CalendarState) => void;
    /** クリアボタン押下 */
    onClear: () => void;
}

const ONE_DAY = 86400000;

/**
 * 日付カレンダー UI 部品。
 * モード切替タブ / 年・月ドロップダウン / 日グリッド / 選択サマリを提供する。
 * Power BI 依存コードは一切持たない純粋 DOM コンポーネント。
 */
export class Calendar {
    private container: HTMLElement;
    private cb: CalendarCallbacks;

    private state: CalendarState = { mode: "single", startDate: "", endDate: "" };
    private viewYear: number;
    private viewMonth: number; // 0-11
    private yearMin = 0;
    private yearMax = 0;
    /** 範囲モードで 1 点目を選んだ状態（まだ終了日未確定） */
    private rangeHalfSet = false;

    // DOM 参照
    private modeTabs: HTMLElement;
    private header: HTMLElement;
    private yearSel: HTMLSelectElement;
    private monthSel: HTMLSelectElement;
    private grid: HTMLElement;
    private summary: HTMLElement;
    private applyBtn: HTMLButtonElement;

    constructor(container: HTMLElement, cb: CalendarCallbacks) {
        this.container = container;
        this.cb = cb;
        const today = new Date();
        this.viewYear = today.getUTCFullYear();
        this.viewMonth = today.getUTCMonth();
        this.setYearRange(this.viewYear - 5, this.viewYear + 5);
        this.build();
    }

    /** データから算出した年レンジを設定（両端 ±2 年のバッファは呼び出し側で加算済み想定） */
    public setYearRange(min: number, max: number): void {
        if (!Number.isFinite(min) || !Number.isFinite(max)) return;
        this.yearMin = Math.min(min, max);
        this.yearMax = Math.max(min, max);
        // viewYear がレンジ外ならクランプ
        if (this.viewYear < this.yearMin) this.viewYear = this.yearMin;
        if (this.viewYear > this.yearMax) this.viewYear = this.yearMax;
        if (this.yearSel) this.fillYearOptions();
    }

    public getState(): CalendarState { return { ...this.state }; }

    public setState(next: Partial<CalendarState>): void {
        const merged: CalendarState = { ...this.state, ...next };
        this.state = merged;
        this.rangeHalfSet = false;
        // view をアクティブな日付に寄せる
        const anchor = merged.startDate || merged.endDate;
        if (anchor) {
            const ep = toDateEpochFromString(anchor);
            if (Number.isFinite(ep)) {
                const d = new Date(ep);
                this.viewYear = d.getUTCFullYear();
                this.viewMonth = d.getUTCMonth();
                // レンジ外なら広げる
                if (this.viewYear < this.yearMin) this.setYearRange(this.viewYear, this.yearMax);
                if (this.viewYear > this.yearMax) this.setYearRange(this.yearMin, this.viewYear);
            }
        }
        this.render();
    }

    private build(): void {
        this.container.classList.add("dc-root");
        while (this.container.firstChild) this.container.removeChild(this.container.firstChild);

        // モードタブ
        this.modeTabs = document.createElement("div");
        this.modeTabs.className = "dc-mode-tabs";
        const singleBtn = this.makeModeBtn("single", "単一");
        const rangeBtn  = this.makeModeBtn("range", "範囲");
        this.modeTabs.appendChild(singleBtn);
        this.modeTabs.appendChild(rangeBtn);
        this.container.appendChild(this.modeTabs);

        // ヘッダー（年・月セレクト + 月送り）
        this.header = document.createElement("div");
        this.header.className = "dc-header";

        this.yearSel = document.createElement("select");
        this.yearSel.className = "dc-year-sel";
        this.yearSel.setAttribute("aria-label", "年");
        this.yearSel.onchange = () => {
            this.viewYear = Number(this.yearSel.value);
            this.render();
        };

        this.monthSel = document.createElement("select");
        this.monthSel.className = "dc-month-sel";
        this.monthSel.setAttribute("aria-label", "月");
        for (let i = 0; i < 12; i++) {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.textContent = MONTH_NAMES_JA[i];
            this.monthSel.appendChild(opt);
        }
        this.monthSel.onchange = () => {
            this.viewMonth = Number(this.monthSel.value);
            this.render();
        };

        const prevBtn = document.createElement("button");
        prevBtn.type = "button";
        prevBtn.className = "dc-nav";
        prevBtn.setAttribute("aria-label", "前月");
        prevBtn.textContent = "◀";
        prevBtn.onclick = () => this.shiftMonth(-1);

        const nextBtn = document.createElement("button");
        nextBtn.type = "button";
        nextBtn.className = "dc-nav";
        nextBtn.setAttribute("aria-label", "翌月");
        nextBtn.textContent = "▶";
        nextBtn.onclick = () => this.shiftMonth(1);

        this.header.appendChild(this.yearSel);
        this.header.appendChild(this.monthSel);
        this.header.appendChild(prevBtn);
        this.header.appendChild(nextBtn);
        this.container.appendChild(this.header);

        // 曜日行
        const wk = document.createElement("div");
        wk.className = "dc-weekdays";
        WEEKDAY_NAMES_JA.forEach(w => {
            const s = document.createElement("span");
            s.className = "dc-weekday";
            s.textContent = w;
            wk.appendChild(s);
        });
        this.container.appendChild(wk);

        // 日グリッド
        this.grid = document.createElement("div");
        this.grid.className = "dc-grid";
        this.container.appendChild(this.grid);

        // サマリ
        this.summary = document.createElement("div");
        this.summary.className = "dc-summary";
        this.container.appendChild(this.summary);

        // フッター（適用ボタン）
        const footer = document.createElement("div");
        footer.className = "dc-footer";
        this.applyBtn = document.createElement("button");
        this.applyBtn.type = "button";
        this.applyBtn.className = "dc-apply-btn";
        this.applyBtn.textContent = "適用";
        this.applyBtn.onclick = () => {
            if (!this.isStateComplete()) return;
            this.cb.onChange(this.getState());
        };
        footer.appendChild(this.applyBtn);
        this.container.appendChild(footer);

        this.fillYearOptions();
        this.render();
    }

    private isStateComplete(): boolean {
        if (this.state.mode === "single") return !!this.state.startDate;
        return !!this.state.startDate && !!this.state.endDate && !this.rangeHalfSet;
    }

    private makeModeBtn(mode: CalendarMode, label: string): HTMLButtonElement {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "dc-mode-btn";
        b.dataset.mode = mode;
        b.textContent = label;
        b.onclick = () => {
            if (this.state.mode === mode) return;
            this.state = { mode, startDate: "", endDate: "" };
            this.rangeHalfSet = false;
            this.render();
        };
        return b;
    }

    private fillYearOptions(): void {
        while (this.yearSel.firstChild) this.yearSel.removeChild(this.yearSel.firstChild);
        for (let y = this.yearMin; y <= this.yearMax; y++) {
            const opt = document.createElement("option");
            opt.value = String(y);
            opt.textContent = `${y}年`;
            this.yearSel.appendChild(opt);
        }
        this.yearSel.value = String(this.viewYear);
    }

    private shiftMonth(delta: number): void {
        let y = this.viewYear;
        let m = this.viewMonth + delta;
        while (m < 0)  { m += 12; y -= 1; }
        while (m > 11) { m -= 12; y += 1; }
        if (y < this.yearMin) this.setYearRange(y, this.yearMax);
        if (y > this.yearMax) this.setYearRange(this.yearMin, y);
        this.viewYear = y;
        this.viewMonth = m;
        this.render();
    }

    private render(): void {
        // モードタブ active
        this.modeTabs.querySelectorAll<HTMLButtonElement>(".dc-mode-btn").forEach(b => {
            b.classList.toggle("active", b.dataset.mode === this.state.mode);
        });

        this.yearSel.value = String(this.viewYear);
        this.monthSel.value = String(this.viewMonth);

        while (this.grid.firstChild) this.grid.removeChild(this.grid.firstChild);

        const firstDow = new Date(Date.UTC(this.viewYear, this.viewMonth, 1)).getUTCDay();
        const daysInMonth = new Date(Date.UTC(this.viewYear, this.viewMonth + 1, 0)).getUTCDate();

        // 先頭余白
        for (let i = 0; i < firstDow; i++) {
            const e = document.createElement("span");
            e.className = "dc-cell dc-empty";
            this.grid.appendChild(e);
        }

        const startEp = this.state.startDate ? toDateEpochFromString(this.state.startDate) : NaN;
        const endEp   = this.state.endDate   ? toDateEpochFromString(this.state.endDate)   : NaN;

        // today はローカル TZ 基準で表示（UI の直感に合わせる）
        const now = new Date();
        const todayStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`;

        for (let d = 1; d <= daysInMonth; d++) {
            const ymd = formatDateUTC(new Date(Date.UTC(this.viewYear, this.viewMonth, d)));
            const ep = toDateEpochFromString(ymd);
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "dc-cell dc-day";
            btn.textContent = String(d);

            if (ymd === todayStr) btn.classList.add("dc-today");

            if (this.state.mode === "single") {
                if (ymd === this.state.startDate) btn.classList.add("dc-selected");
            } else {
                // range
                if (Number.isFinite(startEp) && Number.isFinite(endEp)) {
                    if (ep === startEp)     btn.classList.add("dc-range-start");
                    else if (ep === endEp)  btn.classList.add("dc-range-end");
                    else if (ep > startEp && ep < endEp) btn.classList.add("dc-range-mid");
                } else if (Number.isFinite(startEp) && ep === startEp) {
                    btn.classList.add("dc-range-start");
                }
            }

            btn.onclick = () => this.onDayClick(ymd);
            this.grid.appendChild(btn);
        }

        // サマリ
        while (this.summary.firstChild) this.summary.removeChild(this.summary.firstChild);
        const label = document.createElement("span");
        label.className = "dc-summary-label";
        if (this.state.mode === "single") {
            label.textContent = this.state.startDate || "未選択";
        } else {
            if (this.state.startDate && this.state.endDate) {
                label.textContent = `${this.state.startDate} 〜 ${this.state.endDate}`;
            } else if (this.state.startDate) {
                label.textContent = `${this.state.startDate} 〜 （終了日を選択）`;
            } else {
                label.textContent = "未選択";
            }
        }
        this.summary.appendChild(label);

        if (this.state.startDate || this.state.endDate) {
            const clr = document.createElement("button");
            clr.type = "button";
            clr.className = "dc-clear";
            clr.textContent = "クリア";
            clr.onclick = () => {
                this.state = { mode: this.state.mode, startDate: "", endDate: "" };
                this.rangeHalfSet = false;
                this.render();
                this.cb.onClear();
            };
            this.summary.appendChild(clr);
        }

        if (this.applyBtn) this.applyBtn.disabled = !this.isStateComplete();
    }

    private onDayClick(ymd: string): void {
        if (this.state.mode === "single") {
            this.state = { ...this.state, startDate: ymd, endDate: "" };
            this.render();
            return;
        }
        // range モード
        if (!this.rangeHalfSet) {
            // 1 点目
            this.state = { ...this.state, startDate: ymd, endDate: "" };
            this.rangeHalfSet = true;
            this.render();
            return; // まだ発火しない
        }
        // 2 点目
        let start = this.state.startDate;
        let end = ymd;
        const sEp = toDateEpochFromString(start);
        const eEp = toDateEpochFromString(end);
        if (Number.isFinite(sEp) && Number.isFinite(eEp) && sEp > eEp) {
            [start, end] = [end, start];
        }
        this.state = { mode: "range", startDate: start, endDate: end };
        this.rangeHalfSet = false;
        this.render();
    }
}

export { ONE_DAY };
