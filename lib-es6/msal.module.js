import * as tslib_1 from "tslib";
import { Injectable, NgModule } from "@angular/core";
import { CommonModule } from "@angular/common";
import { MsalService, MSAL_CONFIG } from "./msal.service";
import { MsalGuard } from "./msal-guard.service";
import { BroadcastService } from "./broadcast.service";
Injectable();
var WindowWrapper = /** @class */ (function (_super) {
    tslib_1.__extends(WindowWrapper, _super);
    function WindowWrapper() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return WindowWrapper;
}(Window));
export { WindowWrapper };
var MsalModule = /** @class */ (function () {
    function MsalModule() {
    }
    MsalModule_1 = MsalModule;
    MsalModule.forRoot = function (config) {
        return {
            ngModule: MsalModule_1,
            providers: [
                { provide: MSAL_CONFIG, useValue: config }, MsalService, { provide: WindowWrapper, useValue: window }
            ]
        };
    };
    var MsalModule_1;
    MsalModule = MsalModule_1 = tslib_1.__decorate([
        NgModule({
            imports: [CommonModule],
            declarations: [],
            providers: [MsalGuard, BroadcastService],
        })
    ], MsalModule);
    return MsalModule;
}());
export { MsalModule };
//# sourceMappingURL=msal.module.js.map