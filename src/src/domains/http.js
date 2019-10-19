"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const Utl = require("./utl");
var HTTP_method;
(function (HTTP_method) {
    HTTP_method["GET"] = "GET";
    HTTP_method["POST"] = "POST";
    HTTP_method["PUT"] = "PUT";
    HTTP_method["PATCH"] = "PATCH";
    HTTP_method["DELETE"] = "DELETE";
})(HTTP_method = exports.HTTP_method || (exports.HTTP_method = {}));
var ContentType;
(function (ContentType) {
    ContentType["application_json"] = "application/json";
})(ContentType = exports.ContentType || (exports.ContentType = {}));
var DataType;
(function (DataType) {
    DataType["json"] = "json";
})(DataType = exports.DataType || (exports.DataType = {}));
const timeout_default = 2000;
function create_prms_g(p, timeout_) {
    let timeout = timeout_ ? timeout_ : timeout_default;
    let headers = {};
    if (p.rawToken != null && p.rawToken.length > 0) {
        headers['Authorization'] = 'Bearer ' + p.rawToken;
    }
    if (p.contentType) {
        headers['Content-Type'] = p.contentType;
    }
    if (p.headers) {
        p.headers.forEach((v) => {
            headers[v.name] = v.value;
        });
    }
    let pcf = () => {
        let pr = new Promise((resolve, reject) => {
            $.ajax({
                method: p.method,
                url: p.restUrl,
                data: p.data ? p.data : "",
                dataType: p.dataType ? p.dataType : '',
                headers: headers,
                timeout: timeout
            }).done(function (r) {
                resolve(r);
            }).fail(function (error) {
                reject(error);
            });
        });
        return pr;
    };
    return Utl.retry_promise(pcf);
}
exports.create_prms_g = create_prms_g;
function create_prms(p, timeout_) {
    let timeout = timeout_ ? timeout_ : timeout_default;
    let headers = {};
    if (p.rawToken != null && p.rawToken.length > 0) {
        headers['Authorization'] = 'Bearer ' + p.rawToken;
    }
    if (p.contentType) {
        headers['Content-Type'] = p.contentType;
    }
    if (p.headers) {
        p.headers.forEach((v) => {
            headers[v.name] = v.value;
        });
    }
    let pcf = () => {
        let pr = new Promise((resolve, reject) => {
            $.ajax({
                method: p.method,
                url: p.restUrl,
                data: p.data ? p.data : "",
                dataType: p.dataType ? p.dataType : '',
                headers: headers,
                timeout: timeout
            }).done(function (r) {
                resolve(r);
            }).fail(function (error) {
                reject(error);
            });
        });
        return pr;
    };
    return Utl.retry_promise(pcf);
}
exports.create_prms = create_prms;
function ajax_retry(p) {
    let pcf = () => {
        let pr = $.ajax(p);
        return pr;
    };
    return Utl.retry_promise(pcf);
}
exports.ajax_retry = ajax_retry;
