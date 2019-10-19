"use strict";
//const attemp_num_max = 1;
//const attemp_timeout = 2000;
//export function delay(ms: number) {
//    return new Promise(resolve => setTimeout(resolve, ms));
//}
//export async function retry_promise(cpf: () => Promise<any>) {
//    let error;
Object.defineProperty(exports, "__esModule", { value: true });
//    for (let i = 0; i < attemp_num_max; i++) {
//        let p = cpf();
//        try {
//            if (i == 0) {
//                return await p;
//            }
//            await delay(attemp_timeout);
//            return await p;
//        } catch (err) {
//            error = err;
//        }
//    }
//    throw error;
//}
exports.update = (target, source) => {
    for (var attr in source) {
        if (target.hasOwnProperty(attr))
            target[attr] = source[attr];
    }
};
function xml_decode(s) {
    return s.replace(/&apos;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&gt;/g, '>')
        .replace(/&lt;/g, '<')
        .replace(/&amp;/g, '&')
        .replace(/&#xD;/g, '/r');
}
exports.xml_decode = xml_decode;
function xml_encode(s) {
    return s.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    //.replace(/\r/g, '&#xD;');
}
exports.xml_encode = xml_encode;
function excep_wrap(f, on_ex) {
    try {
        f();
    }
    catch (ex) {
        on_ex(ex);
    }
}
exports.excep_wrap = excep_wrap;
function err_as_string(r) {
    let s = "";
    if (r.status == Office.AsyncResultStatus.Failed) {
        s += (" " + r.status.toString());
        s += (" error name:" + r.error.name);
        s += ("  message:" + r.error.message);
        s += ("  code:" + r.error.code);
    }
    return s;
}
exports.err_as_string = err_as_string;
function any_to_str(o) {
    if (o == null) {
        return "null";
    }
    if (typeof o === "string") {
        return o;
    }
    if (typeof o === 'object') {
        return JSON.stringify(o);
    }
    return o.toString();
}
exports.any_to_str = any_to_str;
function make_err(name, err) {
    let r = { message: "", name: "" };
    if (name) {
        r.name = name;
    }
    r.message = any_to_str(err);
    return r;
}
exports.make_err = make_err;
function wrap_data_error(p) {
    return p.then((d) => ({ data: d }))
        .catch((err) => ({ error: err }));
}
exports.wrap_data_error = wrap_data_error;
function un_wrap_data(p) {
    if (!p.data)
        throw "can't un_wrap_data";
    return p.data;
}
exports.un_wrap_data = un_wrap_data;
function create_prms(fn) {
    return create_prms_3(fn, (v) => { return v; });
}
exports.create_prms = create_prms;
function create_prms_3(fn, rslt_value_convrter) {
    let pr = new Promise((resolve, reject) => {
        fn((r) => {
            let f = () => {
                if (r.status == Office.AsyncResultStatus.Failed) {
                    reject(r.error);
                    return;
                }
                resolve(rslt_value_convrter(r.value));
            };
            excep_wrap(f, (ex) => { reject(ex); });
        });
    });
    return pr;
}
exports.create_prms_3 = create_prms_3;
function addHandlerAsync(mlbx, eventType, handler) {
    return create_prms_3((cbk) => { mlbx.addHandlerAsync(eventType, handler, {}, cbk); }, (v) => { return true; });
}
exports.addHandlerAsync = addHandlerAsync;
//export type Dlg_messageHandler<TRslt> = (arg: I_Dlg_messageHandler_arg, func: (resolve: (value?: TRslt | OfficeExtension.IPromise<TRslt>) => void, reject: (error?: any) => void) => void) => void;
//export type Dlg_eventHandler<TRslt> = (arg: I_Dlg_eventHandler_arg, func: (resolve: (value?: TRslt | OfficeExtension.IPromise<TRslt>) => void, reject: (error?: any) => void) => void) => void;
function error_msg(a) {
    if (!a)
        return null;
    let msg = '';
    a.forEach((v) => {
        if (v.error) {
            msg += any_to_str(v.error);
            msg += " ";
        }
    });
    if (msg == '')
        return null;
    return msg;
}
exports.error_msg = error_msg;
function IsNullOrEmpty(s) {
    if (s == null)
        return true;
    return s.length == 0;
}
exports.IsNullOrEmpty = IsNullOrEmpty;
function IsNullOrEmptyArr(s) {
    if (s == null)
        return true;
    return s.length == 0;
}
exports.IsNullOrEmptyArr = IsNullOrEmptyArr;
//export async function create_prms_2<TRslt>(fn: (cllbck: On_Complete_CallBack) => void):Promise<TRslt> {
//    let pr = new Promise<TRslt>((resolve, reject) => {
//        fn((r: Office.AsyncResult) => {
//            let f = () => {
//                if (r.status == Office.AsyncResultStatus.Failed) {
//                    reject(r.error);
//                    return;
//                }
//                resolve(r.value as TRslt);
//            }
//            excep_wrap(f, (ex) => { reject(ex); })
//        });
//    });
//    return pr;
//}
