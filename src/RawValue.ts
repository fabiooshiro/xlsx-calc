export class RawValue {
    constructor(private value) {

    }
    setValue(v) {
        this.value = v;
    }
    calc() {
        return this.value;
    }
};
