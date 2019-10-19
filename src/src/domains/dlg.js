"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var utl_1 = require("./utl");
var Result;
(function (Result) {
    Result["Yes"] = "Yes";
    Result["No"] = "No";
    Result["Cancel"] = "Cancel";
})(Result = exports.Result || (exports.Result = {}));
;
;
var OffcDlg_wrp = (function () {
    function OffcDlg_wrp() {
    }
    Object.defineProperty(OffcDlg_wrp.prototype, "dialog", {
        get: function () { return this.m_dialog; },
        set: function (dialog) {
            var _this = this;
            this.m_dialog = dialog;
            this.m_dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (a) { _this.on_message(a); });
            this.m_dialog.addEventHandler(Office.EventType.DialogEventReceived, function (a) { _this.on_event(a); });
        },
        enumerable: true,
        configurable: true
    });
    ;
    Object.defineProperty(OffcDlg_wrp.prototype, "on_resolve", {
        get: function () { return this.m_on_resolve; },
        set: function (h) { this.m_on_resolve = h; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(OffcDlg_wrp.prototype, "on_reject", {
        get: function () { return this.m_on_reject; },
        set: function (h) { this.m_on_reject = h; },
        enumerable: true,
        configurable: true
    });
    OffcDlg_wrp.prototype.close = function () {
        this.m_dialog.close();
    };
    return OffcDlg_wrp;
}());
exports.OffcDlg_wrp = OffcDlg_wrp;
function create_dlg_async(dlg_url_relative, options, dlg) {
    return __awaiter(this, void 0, void 0, function () {
        var loc, dlg_url;
        return __generator(this, function (_a) {
            loc = window.location.href.lastIndexOf("/");
            dlg_url = window.location.href.substring(0, loc) + "/" + dlg_url_relative;
            return [2, create_dlg_async_abs(dlg_url, options, dlg)];
        });
    });
}
exports.create_dlg_async = create_dlg_async;
function create_dlg_async_abs(dlg_url, options, dlg) {
    return __awaiter(this, void 0, void 0, function () {
        var fn, pr;
        return __generator(this, function (_a) {
            fn = function (cllbck) {
                Office.context.ui.displayDialogAsync(dlg_url, options, cllbck);
            };
            pr = new Promise(function (resolve, reject) {
                fn(function (r) {
                    var f = function () {
                        if (r.status == Office.AsyncResultStatus.Failed) {
                            reject(r.error);
                            return;
                        }
                        dlg.on_resolve = resolve;
                        dlg.on_reject = reject;
                        dlg.dialog = r.value;
                    };
                    utl_1.excep_wrap(f, function (ex) { reject(ex); });
                });
            });
            return [2, pr];
        });
    });
}
exports.create_dlg_async_abs = create_dlg_async_abs;
var Dlg_YNC;
(function (Dlg_YNC) {
    function on_yes() {
        Office.context.ui.messageParent(Result.Yes);
    }
    Dlg_YNC.on_yes = on_yes;
    function on_no() {
        Office.context.ui.messageParent(Result.No);
    }
    Dlg_YNC.on_no = on_no;
    function on_cancel() {
        Office.context.ui.messageParent(Result.Cancel);
    }
    Dlg_YNC.on_cancel = on_cancel;
    function on_data(d) {
        Office.context.ui.messageParent(d);
    }
    Dlg_YNC.on_data = on_data;
    var Dialog = (function (_super) {
        __extends(Dialog, _super);
        function Dialog(name) {
            var _this = _super.call(this) || this;
            _this.name = "";
            _this.name = name;
            return _this;
        }
        Dialog.prototype.on_message = function (a) {
            switch (a.message) {
                case Result.Yes:
                    this.on_resolve(Result.Yes);
                    break;
                case Result.No:
                    this.on_resolve(Result.No);
                    break;
                default:
                    this.on_resolve(a.message);
                    break;
            }
        };
        Dialog.prototype.on_event = function (a) {
            switch (a.error) {
                case 12002:
                case 12003:
                    this.on_reject(name + " error:" + a.error.toString());
                    break;
                case 12006:
                    this.on_resolve(Result.Cancel);
                    break;
                default:
                    this.on_reject(name + " Undefined error:" + a.error.toString());
                    break;
            }
        };
        return Dialog;
    }(OffcDlg_wrp));
    Dlg_YNC.Dialog = Dialog;
})(Dlg_YNC = exports.Dlg_YNC || (exports.Dlg_YNC = {}));
