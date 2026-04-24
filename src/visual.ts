"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import FilterAction = powerbi.FilterAction;

import {
    AdvancedFilter,
    IAdvancedFilter,
    IFilterColumnTarget,
    FilterType,
    AdvancedFilterConditionOperators,
    AdvancedFilterLogicalOperators,
} from "powerbi-models";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import { Calendar, CalendarMode, CalendarState } from "./calendar";
import {
    formatDateUTC,
    toDateEpoch,
    toDateEpochFromString,
    nextDayEpoch,
    isValidYMD,
} from "./dateUtils";
import { VisualFormattingSettingsModel } from "./settings";

// 全 Date 条件を半開区間 [start, next) で出すための共通 operator
const OP_GTE: AdvancedFilterConditionOperators = "GreaterThanOrEqual";
const OP_LT:  AdvancedFilterConditionOperators = "LessThan";

export class Visual implements IVisual {
    private host: IVisualHost;
    private root: HTMLElement;
    private calEl: HTMLElement;
    private emptyEl: HTMLElement;
    private calendar: Calendar;

    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private lastDataView: DataView | null = null;
    private dateTarget: IFilterColumnTarget | null = null;

    /** 自己発火エコー判定（filterTable 同等パターン） */
    private lastFilterSig = "";

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();

        this.root = document.createElement("div");
        this.root.className = "dc-visual";
        options.element.appendChild(this.root);

        this.emptyEl = document.createElement("div");
        this.emptyEl.className = "dc-empty-msg";
        this.emptyEl.textContent = "日付列を「日付」フィールドにバインドしてください";
        this.root.appendChild(this.emptyEl);

        this.calEl = document.createElement("div");
        this.calEl.className = "dc-calendar";
        this.root.appendChild(this.calEl);

        this.calendar = new Calendar(this.calEl, {
            onChange: (st) => this.onCalendarChange(st),
            onClear: () => this.onCalendarClear(),
        });
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options.dataViews?.[0];
        this.lastDataView = dv ?? null;

        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
        this.applyAppearance();

        const col = dv?.table?.columns?.[0];
        const hasData = !!dv?.table && !!col;
        this.emptyEl.style.display = hasData ? "none" : "";
        this.calEl.style.display = hasData ? "" : "none";
        if (!hasData) return;

        this.dateTarget = this.buildFilterTarget(col);

        // データから年レンジを算出 → カレンダーに反映
        const { min, max } = this.computeYearRange(dv);
        this.calendar.setYearRange(min - 2, max + 2);

        // UI 状態は jsonFilters を唯一の真実源とする。
        // metadata.objects.state は書くが読まない: Power BI の「全フィルターリセット」
        // ブックマークは filter 層のみクリアして metadata は残すため、metadata 読み込みは
        // stale state の温床になる。cross-session 復元は Power BI が PBIX に保存した
        // jsonFilters が担うので metadata 不要。
        this.restoreFromJsonFilters(options.jsonFilters);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ==========================================================
    // カレンダーコールバック
    // ==========================================================

    private onCalendarChange(st: CalendarState): void {
        // 範囲モードで 1 点目しかない場合はまだ発火しない
        if (st.mode === "range" && (!st.startDate || !st.endDate)) return;
        if (st.mode === "single" && !st.startDate) return;
        this.emitFilter(st);
        this.persist(st);
    }

    private onCalendarClear(): void {
        this.emitFilter({ mode: this.calendar.getState().mode, startDate: "", endDate: "" });
        this.persist({ mode: this.calendar.getState().mode, startDate: "", endDate: "" });
    }

    // ==========================================================
    // Filter emit
    // ==========================================================

    private emitFilter(st: CalendarState): void {
        if (!this.dateTarget) return;
        // 値が無ければ remove
        if (!st.startDate && !st.endDate) {
            if (this.lastFilterSig === "") return;
            this.lastFilterSig = "";
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
            return;
        }

        const startYmd = st.mode === "single" ? st.startDate : st.startDate;
        const endYmd   = st.mode === "single" ? st.startDate : st.endDate;
        const startEp = toDateEpochFromString(startYmd);
        const endEp   = toDateEpochFromString(endYmd);
        if (!Number.isFinite(startEp) || !Number.isFinite(endEp)) return;

        const nextEp = nextDayEpoch(endEp);
        const startDate = new Date(startEp);
        const nextDate  = new Date(nextEp);
        const ymdNext   = formatDateUTC(nextDate);

        // 半開区間 [startYmd, ymdNext)
        const conditions = [
            { operator: OP_GTE, value: startDate as unknown as (string | number | boolean) },
            { operator: OP_LT,  value: nextDate  as unknown as (string | number | boolean) },
        ];
        const logical: AdvancedFilterLogicalOperators = "And";
        const filter = new AdvancedFilter(this.dateTarget, logical, ...conditions);

        const sig = `ADV|${this.dateTarget.table}\0${this.dateTarget.column}\0${OP_GTE}:${startYmd}\0${OP_LT}:${ymdNext}`;
        if (sig === this.lastFilterSig) return;
        this.lastFilterSig = sig;
        this.host.applyJsonFilter(filter, "general", "filter", FilterAction.merge);
    }

    // ==========================================================
    // 永続化
    // ==========================================================

    private persist(st: CalendarState): void {
        this.host.persistProperties({
            merge: [{
                objectName: "state",
                selector: null,
                properties: {
                    mode: st.mode,
                    startDate: st.startDate,
                    endDate: st.endDate,
                },
            }],
        });
    }

    private restoreFromPersisted(dv: DataView): boolean {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) return false;
        const mode = (s["mode"] as string) === "range" ? "range" : "single";
        const startDate = String(s["startDate"] ?? "");
        const endDate   = String(s["endDate"]   ?? "");
        if (!isValidYMD(startDate) && !isValidYMD(endDate)) {
            // 空状態として反映（モードのみ）
            this.calendar.setState({ mode: mode as CalendarMode, startDate: "", endDate: "" });
            return true;
        }
        this.calendar.setState({
            mode: mode as CalendarMode,
            startDate: isValidYMD(startDate) ? startDate : "",
            endDate:   isValidYMD(endDate)   ? endDate   : "",
        });
        return true;
    }

    // ==========================================================
    // jsonFilters 受信
    // ==========================================================

    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined): boolean {
        if (!this.dateTarget) return false;

        // AdvancedFilter のみ対象
        const advanced: IAdvancedFilter[] = [];
        if (jsonFilters && jsonFilters.length > 0) {
            for (const f of jsonFilters) {
                const ft = (f as unknown as { filterType?: FilterType })?.filterType;
                if (ft === FilterType.Advanced) advanced.push(f as unknown as IAdvancedFilter);
            }
        }

        // 自分の列に一致するフィルターを探す
        const mine = advanced.find(f => {
            const t = f.target as IFilterColumnTarget;
            return t?.table === this.dateTarget.table && t?.column === this.dateTarget.column;
        });

        // ブックマーク / 外部スライサーで自分の filter が解除された場合は UI をリセット
        if (!mine) {
            if (this.lastFilterSig !== "") {
                this.lastFilterSig = "";
                const currentMode = this.calendar.getState().mode;
                this.calendar.setState({ mode: currentMode, startDate: "", endDate: "" });
            }
            return false;
        }

        const conds = mine.conditions || [];
        // 期待する形: GreaterThanOrEqual <start>, LessThan <next>
        let startEp = NaN;
        let nextEp  = NaN;
        for (const c of conds) {
            const op = c.operator as AdvancedFilterConditionOperators;
            const ep = toDateEpoch(c.value);
            if (!Number.isFinite(ep)) continue;
            if (op === OP_GTE) startEp = ep;
            else if (op === OP_LT) nextEp = ep;
        }
        if (!Number.isFinite(startEp) || !Number.isFinite(nextEp)) return false;
        if (nextEp <= startEp) return false;

        const endEp = nextEp - 86400000;
        const startYmd = formatDateUTC(new Date(startEp));
        const endYmd   = formatDateUTC(new Date(endEp));
        const mode: CalendarMode = startEp === endEp ? "single" : "range";

        // 自己発火エコー判定
        const ymdNext = formatDateUTC(new Date(nextEp));
        const sig = `ADV|${this.dateTarget.table}\0${this.dateTarget.column}\0${OP_GTE}:${startYmd}\0${OP_LT}:${ymdNext}`;
        if (sig === this.lastFilterSig) return true;
        this.lastFilterSig = sig;

        this.calendar.setState({
            mode,
            startDate: startYmd,
            endDate: mode === "range" ? endYmd : "",
        });
        return true;
    }

    // ==========================================================
    // ユーティリティ
    // ==========================================================

    private buildFilterTarget(col: DataViewMetadataColumn): IFilterColumnTarget | null {
        if (!col?.queryName) return null;
        let qn = col.queryName;
        const aggMatch = qn.match(/^\w+\((.+)\)$/);
        const hasAgg = !!aggMatch;
        if (hasAgg) qn = aggMatch[1];
        if (!hasAgg && col.isMeasure) return null;
        const dotIdx = qn.indexOf(".");
        if (dotIdx < 1) return null;
        return { table: qn.substring(0, dotIdx), column: qn.substring(dotIdx + 1) };
    }

    private computeYearRange(dv: DataView): { min: number; max: number } {
        const today = new Date().getUTCFullYear();
        const rows = dv?.table?.rows;
        if (!rows || rows.length === 0) return { min: today - 5, max: today + 5 };
        let min = Number.POSITIVE_INFINITY, max = Number.NEGATIVE_INFINITY;
        for (const r of rows) {
            const v = r?.[0];
            const ep = toDateEpoch(v);
            if (!Number.isFinite(ep)) continue;
            const y = new Date(ep).getUTCFullYear();
            if (y < min) min = y;
            if (y > max) max = y;
        }
        if (!Number.isFinite(min) || !Number.isFinite(max)) {
            return { min: today - 5, max: today + 5 };
        }
        return { min, max };
    }

    private applyAppearance(): void {
        const a = this.formattingSettings?.appearanceCard;
        if (!a) return;
        const s = this.root.style;
        s.setProperty("--dc-font", a.fontFamily.value);
        s.setProperty("--dc-fontsize", `${a.fontSize.value}px`);
        s.setProperty("--dc-accent", a.accentColor.value.value);
        s.setProperty("--dc-fg", a.fontColor.value.value);
        s.setProperty("--dc-bg", a.backgroundColor.value.value);
    }
}
