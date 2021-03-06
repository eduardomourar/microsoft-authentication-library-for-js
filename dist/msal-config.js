export class MsalConfig {
    constructor() {
        this.authority = "https://login.microsoftonline.com/common";
        this.validateAuthority = true;
        this.cacheLocation = "sessionStorage";
        this.loadFrameTimeout = 6000;
        this.navigateToLoginRequestUrl = true;
        this.fallbackToInteractive = false;
    }
}
//# sourceMappingURL=msal-config.js.map