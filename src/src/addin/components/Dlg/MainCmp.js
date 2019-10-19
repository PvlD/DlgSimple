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
var index_1 = require("office-ui-fabric-react/lib/index");
var domains_1 = require("../../../domains");
var MainCmp = (function (_super) {
    __extends(MainCmp, _super);
    function MainCmp(props, context) {
        return _super.call(this, props, context) || this;
    }
    MainCmp.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement("div", { style: { marginLeft: "10px" } },
                React.createElement("div", { style: { height: "10px" } }),
                React.createElement(index_1.PrimaryButton, { autoFocus: true, tabIndex: 1, onClick: function () { Office.context.ui.messageParent(JSON.stringify(domains_1.Dlg.Result.Yes)); }, text: "Yes" }),
                React.createElement("div", { style: { height: "10px" } }),
                React.createElement(index_1.PrimaryButton, { autoFocus: true, tabIndex: 1, onClick: function () { Office.context.ui.messageParent(JSON.stringify(domains_1.Dlg.Result.No)); }, text: "No" }))));
    };
    return MainCmp;
}(React.Component));
exports.MainCmp = MainCmp;
