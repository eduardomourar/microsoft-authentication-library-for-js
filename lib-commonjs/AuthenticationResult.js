"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var AuthenticationResult = /** @class */ (function () {
    function AuthenticationResult(token, tokenType) {
        this._token = "";
        this._tokenType = "";
        this._token = token;
        if (tokenType) {
            this._tokenType = tokenType;
        }
    }
    Object.defineProperty(AuthenticationResult.prototype, "token", {
        get: function () {
            return this._token;
        },
        set: function (value) {
            this._token = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(AuthenticationResult.prototype, "tokenType", {
        get: function () {
            return this._tokenType;
        },
        set: function (value) {
            this._tokenType = value;
        },
        enumerable: true,
        configurable: true
    });
    return AuthenticationResult;
}());
exports.AuthenticationResult = AuthenticationResult;
//# sourceMappingURL=AuthenticationResult.js.map