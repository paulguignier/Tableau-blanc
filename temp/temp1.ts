class DateTimeW {

    private static readonly MIN_EXCEL_DATE = 2;

    public readonly excelValue: number;
    public readonly isRelative: boolean;

    private _computed = false;

    // --- réel ---
    private _year = 0;
    private _month = 0;
    private _day = 0;
    private _hours = 0;
    private _minutes = 0;
    private _seconds = 0;
    private _dayOfWeek?: Day;

    // --- adapté ---
    private _adaptedYear = 0;
    private _adaptedMonth = 0;
    private _adaptedDay = 0;
    private _adaptedHours = 0;

    private constructor(excelValue: number, isRelative: boolean) {
        this.excelValue = excelValue;
        this.isRelative = isRelative;
    }

    public static from(
        value: number | string | DateTime | null | undefined,
        isRelative: boolean = false
    ): DateTime | undefined {

        if (value == null || value === "") return undefined;

        if (value instanceof DateTime) {
            if (value.isRelative !== isRelative) {
                throw new Error(
                    `Un temps ${value.isRelative ? "relatif" : "absolu"} `
                    + `cherche à être affecté à un temps ${isRelative ? "relatif" : "absolu"}.`
                );
            }
            return value;
        }

        const v = Number(value);
        if (!isRelative && v < 0) return undefined;

        return new DateTime(v, isRelative);
    }

    // =============================
    // CALCUL DATE + ADAPTATION
    // =============================
    private compute(): void {
        if (this._computed) return;

        const abs = Math.abs(this.excelValue);
        const dayFraction = abs % 1;

        // ===== date réelle =====
        if (!this.isRelative && this.excelValue > DateTime.MIN_EXCEL_DATE) {
            const base = new Date(Date.UTC(1899, 11, 30));
            const days = Math.floor(this.excelValue);
            const d = new Date(base.getTime() + days * 86400000);

            this._year = d.getUTCFullYear();
            this._month = d.getUTCMonth() + 1;
            this._day = d.getUTCDate();
            this._dayOfWeek = Day.fromNumber(d.getUTCDay());
        }

        // ===== heure réelle =====
        const totalSeconds = Math.round(dayFraction * 86400);
        this._hours = Math.floor(totalSeconds / 3600);
        this._minutes = Math.floor((totalSeconds % 3600) / 60);
        this._seconds = totalSeconds % 60;

        // ===== adaptation ferroviaire =====
        let adaptedDate = { y: this._year, m: this._month, d: this._day };
        let adaptedHours = this._hours;

        if (!this.isRelative && dayFraction < DateTime.rolloverHour && this.excelValue < 2) {

            adaptedHours = this._hours + 24;

            // reculer d’un jour réel
            const d = new Date(Date.UTC(this._year, this._month - 1, this._day));
            d.setUTCDate(d.getUTCDate() - 1);

            adaptedDate = {
                y: d.getUTCFullYear(),
                m: d.getUTCMonth() + 1,
                d: d.getUTCDate()
            };
        }

        this._adaptedYear = adaptedDate.y;
        this._adaptedMonth = adaptedDate.m;
        this._adaptedDay = adaptedDate.d;
        this._adaptedHours = adaptedHours;

        this._computed = true;
    }

    // =============================
    // GETTERS RÉELS
    // =============================
    public get year() { this.compute(); return this._year; }
    public get month() { this.compute(); return this._month; }
    public get day() { this.compute(); return this._day; }
    public get hours() { this.compute(); return this._hours; }
    public get minutes() { this.compute(); return this._minutes; }
    public get seconds() { this.compute(); return this._seconds; }
    public get dayOfWeek() { this.compute(); return this._dayOfWeek; }

    // =============================
    // GETTERS ADAPTÉS
    // =============================
    public get adaptedYear() { this.compute(); return this._adaptedYear; }
    public get adaptedMonth() { this.compute(); return this._adaptedMonth; }
    public get adaptedDay() { this.compute(); return this._adaptedDay; }
    public get adaptedHours() { this.compute(); return this._adaptedHours; }

    // =============================
    // FORMAT
    // =============================
    public format(format: string, adapted: boolean = true): string {

        this.compute();

        const y = adapted ? this._adaptedYear : this._year;
        const m = adapted ? this._adaptedMonth : this._month;
        const d = adapted ? this._adaptedDay : this._day;
        const h = adapted ? this._adaptedHours : this._hours;
        const n = this._minutes;
        const s = this._seconds;

        const pad = (v: number) => v.toString().padStart(2, "0");

        const tokens: Record<string, string> = {
            "yyyy": y.toString(),
            "yy": pad(y % 100),
            "mm": pad(m),
            "m": m.toString(),
            "dd": pad(d),
            "d": d.toString(),
            "hh": pad(h),
            "h": h.toString(),
            "nn": pad(n),
            "n": n.toString(),
            "ss": pad(s),
            "s": s.toString()
        };

        let out = format.toLowerCase();

        Object.keys(tokens)
            .sort((a, b) => b.length - a.length)
            .forEach(t => {
                out = out.replace(new RegExp(t, "g"), tokens[t]);
            });

        return out;
    }

    // =============================
    // OPÉRATIONS TEMPORELLES
    // =============================
    public equalsTo(other?: DateTime): boolean {
        return !!other &&
            this.isRelative === other.isRelative &&
            this.excelValue === other.excelValue;
    }

    public compareTo(other: DateTime): number {
        if (this.isRelative !== other.isRelative) {
            throw new Error("Comparaison relatif / absolu interdite");
        }
        return this.excelValue - other.excelValue;
    }

    public resolveAgainst(reference: DateTime): DateTime {
        if (!this.isRelative) return this;
        if (reference.isRelative) throw new Error("Référence relative");

        return new DateTime(reference.excelValue + this.excelValue, false);
    }

    public relativeTo(reference: DateTime): DateTime {
        if (this.isRelative || reference.isRelative) {
            throw new Error("Deux temps absolus requis");
        }

        return new DateTime(this.excelValue - reference.excelValue, true);
    }

    public add(other: DateTime): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error("Addition seulement relative");
        }
        return new DateTime(this.excelValue + other.excelValue, true);
    }

    public subtract(other: DateTime): DateTime {
        if (!this.isRelative || !other.isRelative) {
            throw new Error("Soustraction seulement relative");
        }
        return new DateTime(this.excelValue - other.excelValue, true);
    }

    public static equalsOrUndefined(a?: DateTime, b?: DateTime): boolean {
        return a === b || (!!a && !!b && a.equalsTo(b));
    }
}