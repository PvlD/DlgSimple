"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var ReactDOM = require("react-dom");
var Fabric_1 = require("office-ui-fabric-react/lib/Fabric");
var MainCmp_1 = require("./MainCmp");
var app_root_id = "app_dlg_dlg";
var fn_init = function () {
    Office.initialize = function (reason) {
        var start = function (f) {
            if (document.readyState != "complete") {
                window.addEventListener("load", f);
            }
            else {
                f();
            }
        };
        start(function () {
            ReactDOM.render(React.createElement(Fabric_1.Fabric, null,
                React.createElement(MainCmp_1.MainCmp, null)), document.getElementById(app_root_id));
        });
    };
};
fn_init();
