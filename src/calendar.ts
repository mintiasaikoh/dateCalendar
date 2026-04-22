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

    // DOM 参照
    private header: HTMLElement;
    private yearSel: HTMLSelectElement;
    private monthSel: HTMLSelectElement;
    private grid: HTMLElement;
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

        // トップバー: 年月セレクト（左）+ 適用ボタン（右）
        const headerBar = document.createElement("div");
        headerBar.className = "dc-header-bar";
        this.container.appendChild(headerBar);

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

        this.header.appendChild(this.yearSel);
        this.header.appendChild(this.monthSel);
        headerBar.appendChild(this.header);

        this.applyBtn = document.createElement("button");
        this.applyBtn.type = "button";
        this.applyBtn.className = "dc-apply-btn";
        this.applyBtn.textContent = "適用";
        this.applyBtn.onclick = () => {
            if (!this.isStateComplete()) return;
            this.cb.onChange(this.getState());
        };
        headerBar.appendChild(this.applyBtn);

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

        this.fillYearOptions();
        this.render();
    }

    private isStateComplete(): boolean {
        return !!this.state.startDate;
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

    private render(): void {
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

            // 単日 or 範囲の表示
            if (this.state.startDate && !this.state.endDate) {
                if (ymd === this.state.startDate) btn.classList.add("dc-selected");
            } else if (Number.isFinite(startEp) && Number.isFinite(endEp)) {
                if (ep === startEp)     btn.classList.add("dc-range-start");
                else if (ep === endEp)  btn.classList.add("dc-range-end");
                else if (ep > startEp && ep < endEp) btn.classList.add("dc-range-mid");
            }

            btn.onclick = () => this.onDayClick(ymd);
            this.grid.appendChild(btn);
        }

        if (this.applyBtn) this.applyBtn.disabled = !this.isStateComplete();
    }

    /**
     * 新仕様:
     *   - 選択なし → クリックで単日選択
     *   - 単日選択中 → 同じ日クリックでクリア、別日クリックで範囲確定
     *   - 範囲選択中 → 同じ start/end クリックでクリア、別日クリックで単日選択からやり直し
     */
    private onDayClick(ymd: string): void {
        const { startDate, endDate } = this.state;

        // クリック日が現状の選択点にヒット → クリア
        if (ymd === startDate || (endDate && ymd === endDate)) {
            this.state = { mode: "single", startDate: "", endDate: "" };
            this.render();
            this.cb.onClear();
            return;
        }

        if (!startDate) {
            // 選択なし → 単日
            this.state = { mode: "single", startDate: ymd, endDate: "" };
            this.render();
            return;
        }

        if (!endDate) {
            // 単日選択中 → 範囲確定
            let s = startDate;
            let e = ymd;
            const sEp = toDateEpochFromString(s);
            const eEp = toDateEpochFromString(e);
            if (Number.isFinite(sEp) && Number.isFinite(eEp) && sEp > eEp) {
                [s, e] = [e, s];
            }
            this.state = { mode: "range", startDate: s, endDate: e };
            this.render();
            return;
        }

        // 範囲選択済み → 新しい単日選択からやり直し
        this.state = { mode: "single", startDate: ymd, endDate: "" };
        this.render();
    }
}

export { ONE_DAY };
