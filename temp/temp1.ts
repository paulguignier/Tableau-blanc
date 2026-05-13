type AssertDDOptions = {
    printSuccess?: boolean;
    printFailure?: boolean;
};

type AssertDDCheck<T> = {
    label: string;
    actual: T | (() => T);
    expected: T | typeof AssertDD.THROWS;
};

class AssertDD {

    /**
     * Constante indiquant qu'une erreur est attendue.
     */
    public static readonly THROWS = Symbol("ASSERT_THROWS");

    /**
     * Nombre total de suites complètes.
     */
    public static completeTests = 0;

    /**
     * Nombre total de suites incomplètes.
     */
    public static incompleteTests = 0;

    private options: AssertDDOptions;

    private total = 0;
    private success = 0;
    private failure = 0;

    /**
     * Constructeur.
     */
    constructor(options: AssertDDOptions = {}) {

        this.options = {
            printSuccess: options.printSuccess ?? true,
            printFailure: options.printFailure ?? true
        };
    }

    /**
     * Réalise plusieurs tests.
     */
    public check<T>(
        checks: AssertDDCheck<T>[],
        options: AssertDDOptions = {}
    ): boolean {

        const printSuccess =
            options.printSuccess ?? this.options.printSuccess;

        const printFailure =
            options.printFailure ?? this.options.printFailure;

        let globalSuccess = true;

        checks.forEach(({label, actual, expected}) => {

            let actualValue: T | undefined;
            let ok = false;

            try {

                actualValue =
                    typeof actual === "function"
                        ? (actual as () => T)()
                        : actual;

                ok =
                    expected !== AssertDD.THROWS
                    && actualValue === expected;

            } catch (error) {

                ok = expected === AssertDD.THROWS;

                actualValue = error as T;
            }

            this.total++;

            if (ok) {
                this.success++;
            } else {
                this.failure++;
                globalSuccess = false;
            }

            if (ok) {

                if (printSuccess) {

                    CONSOLE.log(
                        `✔ ${label}`
                    );
                }

            } else {

                if (printFailure) {

                    if (expected === AssertDD.THROWS) {

                        CONSOLE.log(
                            `✘ ${label} | erreur attendue mais aucune erreur levée`
                        );

                    } else {

                        CONSOLE.log(
                            `✘ ${label}`
                            + ` | attendu: ${expected}`
                            + ` | obtenu: ${actualValue}`
                        );
                    }
                }
            }
        });

        return globalSuccess;
    }

    /**
     * Imprime le résumé.
     */
    public printSummary(
        title: string = "Résultats des tests",
        reset: boolean = true
    ): void {

        const complete = this.failure === 0;

        if (complete) {
            AssertDD.completeTests++;
        } else {
            AssertDD.incompleteTests++;
        }

        CONSOLE.log(
            `${title} : `
            + `${this.success} / ${this.total} réussis`
            + ` (échecs : ${this.failure})`
        );

        CONSOLE.log(
            `Suites globales : `
            + `${AssertDD.completeTests} complètes`
            + ` | ${AssertDD.incompleteTests} incomplètes`
        );

        if (reset) {
            this.reset();
        }
    }

    /**
     * Réinitialise les compteurs.
     */
    public reset(): void {

        this.total = 0;
        this.success = 0;
        this.failure = 0;
    }
}