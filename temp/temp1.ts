function testTrainNumber(options: Partial<AssertDDOptions> = {}) {

    const assert = new AssertDD(options);
    TrainNumber.load(true);

    // ------------------------------------------------------------
    // Constructeur
    // ------------------------------------------------------------

    const constructorTests = [
        { desc: "Nombre simple", input: 146490, expected: "146490" },
        { desc: "Chaîne avec slash", input: "146490/1", expected: "146490/146491" },
        { desc: "Minuscules + parasites", input: "w-14a6490", expected: "W14A6490" }
    ];

    constructorTests.forEach(t => {
        const tn = new TrainNumber(t.input);
        assert.check(
            `new TrainNumber(${JSON.stringify(t.input)}) (${t.desc})`,
            tn.doubleParity ? tn.value : tn.value,
            t.expected
        );
    });

    // ------------------------------------------------------------
    // isW()
    // ------------------------------------------------------------

    const wTests = [
        { value: 146490, expected: true },
        { value: 569907, expected: true },
        { value: 147490, expected: false },
        { value: 165470, expected: false }
    ];

    wTests.forEach(t => {
        const tn = new TrainNumber(t.value);
        assert.check(`isW(${t.value})`, tn.isW(), t.expected);
    });

    // ------------------------------------------------------------
    // abbreviateTo4Digits()
    // ------------------------------------------------------------

    const abbreviateTests = [
        { value: 146490, abbreviate: true, expected: "6490" },
        { value: 569907, abbreviate: true, expected: "569907" },
        { value: 146490, abbreviate: false, expected: "146490" }
    ];

    abbreviateTests.forEach(t => {
        const tn = new TrainNumber(t.value);
        assert.check(
            `abbreviate(${t.value})`,
            tn.abbreviateTo4Digits(t.abbreviate),
            t.expected
        );
    });

    // ------------------------------------------------------------
    // adaptWithParity()
    // ------------------------------------------------------------

    const parityTests = [
        { value: 146491, parity: Parity.even, expected: "146490" },
        { value: 146490, parity: Parity.odd, expected: "146491" },
        { value: 146490, parity: Parity.double, expected: "146490/146491" },
        { value: 146490, parity: Parity.double, abbreviate: true, expected: "6490/6491" }
    ];

    parityTests.forEach(t => {
        const tn = new TrainNumber(t.value);
        assert.check(
            `adaptWithParity(${t.value}, ${t.parity})`,
            tn.adaptWithParity(t.parity, t.abbreviate),
            t.expected
        );
    });

    assert.printSummary("testTrainNumber");
}
