"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var ReactDOM = require("react-dom");
var MainCmp_1 = require("./components/MainCmp");
var app_root_id = "dlg_app";
var fn_init = function () {
    Office.onReady().then(function (reason) {
        var ff = function () {
            ReactDOM.render(React.createElement(MainCmp_1.MainCmp, null), document.getElementById(app_root_id));
        };
        ff();
    });
};
fn_init();
