private static optimiseCode(numbersString: string): string {

    let remaining = numbersString;
    let result = '';

    const rules = [...this.compressionRules]
        .sort((a, b) => b.numbers.length - a.numbers.length);

        for (const r of rules) {
            if (r.numbers.split('').every(n => remaining.includes(n))) {
                result += r.code;
                r.numbers.split('').forEach(n => {
                    remaining = remaining.replace(n, '');
                });
            }
        }
    
        return result + remaining;
}