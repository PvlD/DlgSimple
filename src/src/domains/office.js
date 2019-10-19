"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const Utl = require("../domains/utl");
const app_1 = require("../app/app");
// removeAsync
function statusClean(key, report_err, on_done) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, (r) => {
        if (r.status == Office.AsyncResultStatus.Failed) {
            if (report_err) {
                app_1.Log.error(" statusClean  " + Utl.err_as_string(r));
            }
        }
        if (on_done) {
            on_done();
        }
    });
}
exports.statusClean = statusClean;
function statusUpdate(icon, text, key, persistent, type, on_done) {
    let JSONmessage = {
        type: type,
        message: text,
    };
    if (icon) {
        JSONmessage.icon = icon;
    }
    if (type == Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
        JSONmessage.persistent = persistent;
    }
    Office.context.mailbox.item.notificationMessages.replaceAsync(key, JSONmessage, {}, (r) => {
        if (r.status == Office.AsyncResultStatus.Failed) {
            app_1.Log.error(" statusUpdate  " + Utl.err_as_string(r));
        }
        if (on_done) {
            on_done();
        }
    });
}
exports.statusUpdate = statusUpdate;
