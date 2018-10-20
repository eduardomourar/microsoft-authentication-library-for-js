"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var MSALError = /** @class */ (function () {
    function MSALError(error, errorDesc, scopes) {
        this._error = "";
        this._errorDesc = "";
        this._scopes = "";
        this._error = error;
        if (errorDesc) {
            this._errorDesc = errorDesc;
        }
        if (scopes) {
            this._scopes = scopes;
        }
    }
    Object.defineProperty(MSALError.prototype, "error", {
        get: function () {
            return this._error;
        },
        set: function (value) {
            this._error = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(MSALError.prototype, "errorDesc", {
        get: function () {
            return this._errorDesc;
        },
        set: function (value) {
            this._errorDesc = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(MSALError.prototype, "scopes", {
        get: function () {
            return this._scopes;
        },
        set: function (value) {
            this._scopes = value;
        },
        enumerable: true,
        configurable: true
    });
    return MSALError;
}());
exports.MSALError = MSALError;
//# sourceMappingURL=MSALError.js.map