"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const Utl = require("./utl");
function getAccessTokenAsync(options) {
    return Utl.create_prms((cbk) => { Office.context.auth.getAccessTokenAsync((options == null ? null : options), cbk); });
}
exports.getAccessTokenAsync = getAccessTokenAsync;
