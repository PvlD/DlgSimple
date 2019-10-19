"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var ReactDOM = require("react-dom");
var react_redux_1 = require("react-redux");
var index_1 = require("./store/index");
var main_1 = require("./components/main");
var actions_1 = require("./actions");
var RST = require("./domains/rest");
var app_1 = require("./app/app");
var office_1 = require("./domains/office");
var app_root_id = "mw_app";
var fn_start_processing = null;
var fn_init = function () {
    Office.initialize = function (reason) {
        RST.init();
        var start = function (f) {
            if (document.readyState != "complete") {
                window.addEventListener("load", f);
            }
            else {
                f();
            }
        };
        start(function () {
            //debugger;
            var store = index_1.configureStore();
            fn_start_processing = function (clbk) { store.dispatch(actions_1.OnSendActions.start_processing({ data: clbk })); };
            ReactDOM.render(React.createElement(react_redux_1.Provider, { store: store },
                React.createElement(main_1.MainCmp, null)), document.getElementById(app_root_id));
            //Log.error("after render MainCmp");
            store.dispatch(actions_1.OnSendActions.init({ url: "." }));
            //store.dispatch(OnSendActions.test_ews({ data:"" }));
        });
    };
};
fn_init();
window.mw_on_send = function (event) {
    //Log.error("mw_on_send start ");
    var clbk = new office_1.AppCommandCallback(event);
    if (!fn_start_processing) {
        app_1.Log.error("!fn_start_processing");
        clbk.cancel();
        return;
    }
    fn_start_processing(clbk);
    //Log.error("mw_on_send");
};
