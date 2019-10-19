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
var React = require("react");
var DlgSimple = require("../components/dlg/dlg_wrp");
var Dlg = require("../../domains/dlg");
var index_1 = require("office-ui-fabric-react/lib/index");
var MainCmp = (function (_super) {
    __extends(MainCmp, _super);
    function MainCmp(props, context) {
        return _super.call(this, props, context) || this;
    }
    MainCmp.prototype.show_dlg = function () {
        var dlg = new DlgSimple.Dialog();
        var dlg_url = "./DlgSimple.html";
        var dlgp = Dlg.create_dlg_async(dlg_url, { height: 40, width: 50, displayInIframe: false, promptBeforeOpen: true }, dlg);
        dlgp
            .then(function (d) {
            var dd = JSON.parse(d);
            switch (dd) {
                case Dlg.Result.Cancel:
                    {
                        console.info("Cancel");
                    }
                    break;
                case Dlg.Result.Yes:
                    {
                        console.info("Yes");
                    }
                    break;
                case Dlg.Result.No:
                    {
                        console.info("No");
                    }
                    break;
                default:
                    {
                        console.error("Unknown Dlg.Result");
                    }
            }
            dlg.close();
        })
            .catch(function (err) {
            dlg.close();
            console.error(err);
        });
    };
    MainCmp.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", null,
            React.createElement("div", { style: { height: "10px" } }),
            React.createElement(index_1.PrimaryButton, { autoFocus: true, tabIndex: 1, onClick: function () { _this.show_dlg(); }, text: "Show Dialog" })));
    };
    return MainCmp;
}(React.Component));
exports.MainCmp = MainCmp;
