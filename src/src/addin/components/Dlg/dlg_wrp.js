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
Object.defineProperty(exports, "__esModule", { value: true });
var dialog_name = "DlgSimple";
var dlg_1 = require("../../../domains/dlg");
exports.Dlg_YNC = dlg_1.Dlg_YNC;
var Dialog = (function (_super) {
    __extends(Dialog, _super);
    function Dialog() {
        return _super.call(this, dialog_name) || this;
    }
    return Dialog;
}(dlg_1.Dlg_YNC.Dialog));
exports.Dialog = Dialog;
