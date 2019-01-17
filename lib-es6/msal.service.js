import * as tslib_1 from "tslib";
import { Inject, Injectable, InjectionToken } from "@angular/core";
import { MsalConfig } from "./msal-config";
import "rxjs/add/observable/of";
import "rxjs/add/operator/do";
import "rxjs/add/operator/delay";
import { UserAgentApplication, Constants, Logger } from "msal";
import { Router } from "@angular/router";
import { BroadcastService } from "./broadcast.service";
import { AuthenticationResult } from "./AuthenticationResult";
import { MSALError } from "./MSALError";
export var MSAL_CONFIG = new InjectionToken("MSAL_CONFIG");
var MsalService = /** @class */ (function (_super) {
    tslib_1.__extends(MsalService, _super);
    function MsalService(config, router, broadcastService) {
        var _this = _super.call(this, config.clientID, config.authority, null, {
            validateAuthority: config.validateAuthority,
            cacheLocation: config.cacheLocation,
            redirectUri: config.redirectUri,
            postLogoutRedirectUri: config.postLogoutRedirectUri,
            logger: new Logger(config.logger, { correlationId: config.correlationId, level: config.level, piiLoggingEnabled: config.piiLoggingEnabled }),
            loadFrameTimeout: config.loadFrameTimeout,
            navigateToLoginRequestUrl: config.navigateToLoginRequestUrl,
            isAngular: true,
            unprotectedResources: config.unprotectedResources,
            protectedResourceMap: new Map(config.protectedResourceMap),
        }) || this;
        _this.config = config;
        _this.router = router;
        _this.broadcastService = broadcastService;
        _this._oauthData = { isAuthenticated: false, userName: "", loginError: "", idToken: {} };
        _this.loginScopes = [_this.clientId];
        _this.updateDataFromCache(_this.loginScopes);
        var urlHash = window.location.hash;
        _this.processHash(urlHash);
        window.addEventListener("msal:popUpHashChanged", function (e) {
            _this._logger.verbose("popUpHashChanged ");
            _this.processHash(e.detail);
        });
        window.addEventListener('msal:popUpClosed', function (e) {
            var errorParts = e.detail.split('|');
            var msalError = new MSALError(errorParts[0], errorParts[1]);
            if (_this.loginInProgress()) {
                broadcastService.broadcast('msal:loginFailure', msalError);
                _this.setloginInProgress(false);
            }
            else if (_this.getAcquireTokenInProgress()) {
                broadcastService.broadcast('msal:acquireTokenFailure', msalError);
                _this.setAcquireTokenInProgress(false);
            }
        });
        _this.router.events.subscribe(function (event) {
            for (var i = 0; i < router.config.length; i++) {
                if (!router.config[i].canActivate) {
                    if (_this.config && _this.config.unprotectedResources) {
                        if (!_this.isUnprotectedResource(router.config[i].path) && !_this.isEmpty(router.config[i].path)) {
                            _this.config.unprotectedResources.push(router.config[i].path);
                        }
                    }
                }
            }
        });
        return _this;
    }
    MsalService.prototype.updateDataFromCache = function (scopes) {
        // only cache lookup here to not interrupt with events
        var cacheResult;
        cacheResult = _super.prototype.getCachedTokenInternal.call(this, scopes, this.getUser());
        this._oauthData.isAuthenticated = cacheResult != null && cacheResult.token !== null && cacheResult.token.length > 0;
        var user = this.getUser();
        if (user) {
            this._oauthData.userName = user.name;
            this._oauthData.idToken = user.idToken;
        }
        if (cacheResult && cacheResult.error) {
            this._oauthData.loginError = cacheResult == null ? "" : cacheResult.error;
        }
    };
    MsalService.prototype.processHash = function (hash) {
        var _this = this;
        if (this.isCallback(hash)) {
            var isPopup = false;
            var requestInfo = null;
            var callback = null;
            var msal;
            // callback can come from popupWindow, iframe or mainWindow
            if (window.openedWindows.length > 0 && window.openedWindows[window.openedWindows.length - 1].opener
                && window.openedWindows[window.openedWindows.length - 1].opener.msal) {
                var mainWindow = window.openedWindows[window.openedWindows.length - 1].opener;
                msal = mainWindow.msal;
                isPopup = true;
                requestInfo = msal.getRequestInfo(hash);
                if (mainWindow.callBackMappedToRenewStates[requestInfo.stateResponse]) {
                    callback = mainWindow.callBackMappedToRenewStates[requestInfo.stateResponse];
                }
            }
            //redirect flow
            else if (window.parent && window.parent.msal) {
                msal = window.parent.msal;
                requestInfo = msal.getRequestInfo(hash);
                if (window.parent !== window && window.parent.callBackMappedToRenewStates[requestInfo.stateResponse]) {
                    callback = window.parent.callBackMappedToRenewStates[requestInfo.stateResponse];
                }
                else {
                    callback = msal._tokenReceivedCallback;
                }
            }
            this.getLogger().verbose("Processing the hash: " + hash);
            this.saveTokenFromHash(requestInfo);
            // Return to callback if it is sent from iframe
            var token = requestInfo.parameters["access_token"] || requestInfo.parameters["id_token"];
            var error = requestInfo.parameters["error"];
            var errorDescription = requestInfo.parameters["error_description"];
            var tokenType = null;
            var msalError = new MSALError(error, errorDescription);
            var authenticationResult = new AuthenticationResult(token);
            if (requestInfo.stateMatch) {
                if (requestInfo.requestType === "RENEW_TOKEN") {
                    tokenType = Constants.accessToken;
                    authenticationResult.tokenType = tokenType;
                    this._renewActive = false;
                    // Call within the same context without full page redirect keeps the callback
                    // id_token or access_token can be renewed
                    if (window.parent === window && !window.parent.callBackMappedToRenewStates[requestInfo.stateResponse]) {
                        if (token) {
                            this.broadcastService.broadcast("msal:acquireTokenSuccess", authenticationResult);
                        }
                        else if (error && errorDescription) {
                            this.broadcastService.broadcast("msal:acquireTokenFailure", msalError);
                        }
                    }
                }
                else if (requestInfo.requestType === "LOGIN") {
                    tokenType = Constants.idToken;
                    authenticationResult.tokenType = tokenType;
                    this.updateDataFromCache(this.loginScopes);
                    if (this._oauthData.userName) {
                        setTimeout(function () {
                            // id_token is added as token for the app
                            _this.updateDataFromCache(_this.loginScopes);
                            //todo temp commented
                            //  this.userInfo = this._oauthData;
                        }, 1);
                        this.broadcastService.broadcast("msal:loginSuccess", authenticationResult);
                    }
                    else {
                        this.broadcastService.broadcast("msal:loginFailure", msalError);
                    }
                }
                if (callback && typeof callback === "function") {
                    callback(errorDescription, token, error, tokenType);
                }
                // since this is a token renewal request in iFrame, we don't need to proceed with the location change.
                if (window.parent !== window) { //in iframe
                    if (event && event.preventDefault) {
                        event.preventDefault();
                    }
                    return;
                }
                // redirect to login start page
                if (window.parent === window && !isPopup) {
                    if (this._navigateToLoginRequestUrl) {
                        var loginStartPage = this._cacheStorage.getItem(Constants.loginRequest);
                        this._cacheStorage.setItem(Constants.urlHash, hash);
                        if (typeof loginStartPage !== "undefined" && loginStartPage && loginStartPage.length !== 0) {
                            // prevent the current location change and redirect the user back to the login start page
                            this._logger.verbose("Redirecting to start page: " + loginStartPage);
                            window.location.href = loginStartPage;
                        }
                    }
                    //redirect to redirect uri. No page reload here since we are only removing the url after the hash
                    else {
                        window.location.hash = '';
                    }
                }
            }
            else {
                // state did not match, broadcast an error
                this.broadcastService.broadcast("msal:stateMismatch", msalError);
            }
        }
        else {
            var pendingCallback = this._cacheStorage.getItem(Constants.urlHash);
            if (pendingCallback) {
                this.processRedirectCallBack(pendingCallback);
            }
        }
    };
    ;
    MsalService.prototype.processRedirectCallBack = function (hash) {
        this._logger.info('Processing the callback from redirect response');
        var requestInfo = this.getRequestInfo(hash);
        var token = requestInfo.parameters[Constants.accessToken] || requestInfo.parameters[Constants.idToken];
        var errorDesc = requestInfo.parameters[Constants.errorDescription];
        var error = requestInfo.parameters[Constants.error];
        var tokenType;
        this._cacheStorage.removeItem(Constants.urlHash);
        var msalError = new MSALError(error, errorDesc);
        var authenticationResult = new AuthenticationResult(token);
        if (requestInfo.parameters[Constants.accessToken]) {
            tokenType = Constants.accessToken;
            if (token) {
                authenticationResult.tokenType = tokenType;
                this.broadcastService.broadcast("msal:acquireTokenSuccess", authenticationResult);
            }
            else if (error && errorDesc) {
                //TODO this should also send back the scopes
                this.broadcastService.broadcast("msal:acquireTokenFailure", msalError);
            }
        }
        else {
            tokenType = Constants.idToken;
            if (token) {
                authenticationResult.tokenType = tokenType;
                this.broadcastService.broadcast("msal:loginSuccess", authenticationResult);
            }
            else if (error && errorDesc) {
                this.broadcastService.broadcast("msal:loginFailure", msalError);
            }
        }
    };
    MsalService.prototype.isUnprotectedResource = function (url) {
        if (this.config && this.config.unprotectedResources) {
            for (var i = 0; i < this.config.unprotectedResources.length; i++) {
                if (url.indexOf(this.config.unprotectedResources[i]) > -1) {
                    return true;
                }
            }
        }
        return false;
    };
    MsalService.prototype.isEmpty = function (str) {
        return (typeof str === "undefined" || !str || 0 === str.length);
    };
    //dummy method for future use
    MsalService.prototype.authCallback = function (errorDesc, _token, error, _tokenType) {
    };
    MsalService.prototype.clearCache = function () {
        _super.prototype.clearCache.call(this);
    };
    /*This is a private api and not supposed to be use by customers */
    MsalService.prototype.getLogger = function () {
        return _super.prototype.getLogger.call(this);
    };
    MsalService.prototype.getCacheStorage = function () {
        return this._cacheStorage;
    };
    MsalService.prototype.isCallback = function (hash) {
        return _super.prototype.isCallback.call(this, hash);
    };
    MsalService.prototype.loginRedirect = function (consentScopes, extraQueryParameters) {
        this._logger.verbose("login redirect flow");
        _super.prototype.loginRedirect.call(this, consentScopes, extraQueryParameters);
    };
    MsalService.prototype.loginPopup = function (consentScopes, extraQueryParameters) {
        var _this = this;
        this._logger.verbose("login popup flow");
        return new Promise(function (resolve, reject) {
            _super.prototype.loginPopup.call(_this, consentScopes, extraQueryParameters).then(function (idToken) {
                var authenticationResult = new AuthenticationResult(idToken, "idToken");
                _this.broadcastService.broadcast("msal:loginSuccess", authenticationResult);
                resolve(idToken);
            }, function (error) {
                var errorParts = error.split('|');
                var msalError = new MSALError(errorParts[0], errorParts[1]);
                _this._logger.error("Error during login:\n" + error);
                _this.broadcastService.broadcast("msal:loginFailure", msalError);
                reject(error);
            });
        });
    };
    MsalService.prototype.logout = function () {
        this.user = null;
        _super.prototype.logout.call(this);
    };
    MsalService.prototype.getCachedTokenInternal = function (scopes) {
        return _super.prototype.getCachedTokenInternal.call(this, scopes, this.getUser());
    };
    MsalService.prototype.acquireTokenSilent = function (scopes, authority, user, extraQueryParameters) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _super.prototype.acquireTokenSilent.call(_this, scopes, authority, user, extraQueryParameters)
                .then(function (token) {
                _this._renewActive = false;
                var authenticationResult = new AuthenticationResult(token);
                _this.broadcastService.broadcast('msal:acquireTokenSuccess', authenticationResult);
                resolve(token);
            }, function (error) {
                if (error === void 0) { error = ''; }
                if (_this.config && _this.config.fallbackToInteractive) {
                    if (_this.config.popUp) {
                        return _this.acquireTokenPopup(scopes, authority, user, extraQueryParameters);
                    }
                    else {
                        try {
                            _this.acquireTokenRedirect(scopes, authority, user, extraQueryParameters);
                            return resolve(true);
                        }
                        catch (err) {
                            error += ' ' + err;
                        }
                    }
                }
                var errorParts = error.split('|');
                var msalError = new MSALError(errorParts[0], errorParts[1]);
                _this._renewActive = false;
                _this.broadcastService.broadcast('msal:acquireTokenFailure', msalError);
                _this._logger.error('Error when acquiring token for scopes: ' + scopes + " " + error);
                reject(error);
            });
        });
    };
    MsalService.prototype.acquireTokenPopup = function (scopes, authority, user, extraQueryParameters) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _super.prototype.acquireTokenPopup.call(_this, scopes, authority, user, extraQueryParameters).then(function (token) {
                _this._renewActive = false;
                var authenticationResult = new AuthenticationResult(token);
                _this.broadcastService.broadcast('msal:acquireTokenSuccess', authenticationResult);
                resolve(token);
            }, function (error) {
                var errorParts = error.split('|');
                var msalError = new MSALError(errorParts[0], errorParts[1]);
                _this._renewActive = false;
                _this.broadcastService.broadcast('msal:acquireTokenFailure', msalError);
                _this._logger.error('Error when acquiring token for scopes : ' + scopes + " " + error);
                reject(error);
            });
        });
    };
    MsalService.prototype.acquireTokenRedirect = function (scopes, authority, user, extraQueryParameters) {
        var acquireTokenStartPage = this._cacheStorage.getItem(Constants.loginRequest);
        if (window.location.href !== acquireTokenStartPage)
            this._cacheStorage.setItem(Constants.loginRequest, window.location.href);
        _super.prototype.acquireTokenRedirect.call(this, scopes, authority, user, extraQueryParameters);
    };
    MsalService.prototype.acquireTokenInProgress = function () {
        return _super.prototype.getAcquireTokenInProgress.call(this);
    };
    MsalService.prototype.loginInProgress = function () {
        return _super.prototype.loginInProgress.call(this);
    };
    MsalService.prototype.getUser = function () {
        return _super.prototype.getUser.call(this);
    };
    MsalService.prototype.getScopesForEndpoint = function (endpoint) {
        return _super.prototype.getScopesForEndpoint.call(this, endpoint);
    };
    MsalService.prototype.clearCacheForScope = function (accessToken) {
        _super.prototype.clearCacheForScope.call(this, accessToken);
    };
    MsalService.prototype.info = function (message) {
        this._logger.info(message);
    };
    MsalService.prototype.verbose = function (message) {
        this._logger.verbose(message);
    };
    MsalService.prototype.removeItem = function (key) {
        this._cacheStorage.removeItem(key);
    };
    MsalService = tslib_1.__decorate([
        Injectable(),
        tslib_1.__param(0, Inject(MSAL_CONFIG)),
        tslib_1.__metadata("design:paramtypes", [MsalConfig, Router, BroadcastService])
    ], MsalService);
    return MsalService;
}(UserAgentApplication));
export { MsalService };
//# sourceMappingURL=msal.service.js.map