"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const React = require("react");
const react_redux_1 = require("react-redux");
const redux_1 = require("redux");
const DlgDropBox = require("../components/DlgDropBox/dlg_wrp");
const Dlg = require("../../domains/dlg");
const index_1 = require("office-ui-fabric-react/lib/index");
const index_2 = require("office-ui-fabric-react/lib/index");
const DetailsList_1 = require("office-ui-fabric-react/lib/DetailsList");
const MarqueeSelection_1 = require("office-ui-fabric-react/lib/MarqueeSelection");
const Icons_1 = require("office-ui-fabric-react/lib/Icons");
Icons_1.initializeIcons( /* optional base url */);
const Utl = require("../../domains/utl");
const utils_1 = require("../../utils");
const actions_1 = require("../actions");
const App = require("../../app/app");
const HTTP = require("../../domains/http");
const domains_1 = require("../../domains");
const models_1 = require("../models");
const Api = require("../api/api");
function getRestId(ewsId) {
    return Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
}
// EML = 1, MSG = 2 , PDF=3, ZIP=4 
const INITIAL_OPTIONS_conversion_type = [
    { key: 'MSG', text: 'MSG' },
    { key: 'EML', text: 'EML' },
    { key: 'PDF', text: 'PDF' },
    { key: 'ZIP', text: 'ZIP' }
];
const INITIAL_OPTIONS_SaveTo = [
    { key: models_1.Cmd.SaveTo[models_1.Cmd.SaveTo.OneDtrive], text: models_1.Cmd.SaveTo[models_1.Cmd.SaveTo.OneDtrive] },
    { key: models_1.Cmd.SaveTo[models_1.Cmd.SaveTo.DropBox], text: models_1.Cmd.SaveTo[models_1.Cmd.SaveTo.DropBox] },
];
const INITIAL_OPTIONS_RootItemConversion = [
    { key: models_1.Cmd.RootItemConversion[models_1.Cmd.RootItemConversion.PDF], text: models_1.Cmd.RootItemConversion[models_1.Cmd.RootItemConversion.PDF] },
    { key: models_1.Cmd.RootItemConversion[models_1.Cmd.RootItemConversion.HTML], text: models_1.Cmd.RootItemConversion[models_1.Cmd.RootItemConversion.HTML] },
];
// AttachedItemConversion
const INITIAL_OPTIONS_AttachedItemConversion = [
    { key: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.EML], text: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.EML] },
    { key: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.MSG], text: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.MSG] },
    { key: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.PDF], text: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.PDF] },
    { key: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.ZIP], text: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.ZIP] },
];
;
class MainCmp_ extends React.Component {
    constructor(props, context) {
        super(props, context);
        this._selection_att = new DetailsList_1.Selection({
            onSelectionChanged: () => this._on_Selection_att_changed()
        });
        this._columns_att = [
            { key: 'column1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
        ];
        this.state = {
            tk: "",
            info: "",
            attch: {
                data: Office.context.mailbox.item.attachments.filter((e, i) => {
                    e.mw_i = i;
                    return !e.isInline;
                }),
                is_item_present: (Office.context.mailbox.item.attachments.findIndex((e) => { return e.attachmentType == Office.MailboxEnums.AttachmentType.Item; }) != -1)
            },
            cmd: {
                SaveTo: models_1.Cmd.SaveTo.OneDtrive,
                SaveAs: models_1.Cmd.SaveAs.MSG,
                ItmId: getRestId(Office.context.mailbox.item.itemId),
                AttachedItemConversion: models_1.Cmd.AttachedItemConversion.MSG,
                AttchmentList: [],
                AttchmentProcessingKind: models_1.Cmd.AttchmentProcessingKind.All,
                RootItemConversion: models_1.Cmd.RootItemConversion.PDF,
            },
            drop_box: {
                enabled: false,
                is_log_in: false,
                tkn: "",
            }
        };
        this.on_get_dp_status();
    }
    _on_Selection_att_changed() {
        const selectionCount = this._selection_att.getSelectedCount();
        const attchmentProcessingKind = this._selection_att.isAllSelected() ? models_1.Cmd.AttchmentProcessingKind.All : selectionCount == 0 ? models_1.Cmd.AttchmentProcessingKind.None : models_1.Cmd.AttchmentProcessingKind.List;
        let attchmentList = [];
        if (attchmentProcessingKind == models_1.Cmd.AttchmentProcessingKind.List) {
            attchmentList = this._selection_att.getSelection().map((e) => {
                return e.mw_i;
                //return (e as Office.AttachmentDetails).id;
            });
        }
        this.setState((ps) => (Object.assign({}, ps, { cmd: Object.assign({}, ps.cmd, { AttchmentProcessingKind: attchmentProcessingKind, AttchmentList: attchmentList }) })));
    }
    set_info(d) {
        this.setState((ps) => ({ info: d }));
    }
    get_url_ses() {
        let url = window.location.protocol + "//" + window.location.hostname + "/" + "Messageware-SES";
        return url;
    }
    //private drop_box_save_token(t:string,code: string) {
    //    let base_url = this.get_url_ses();
    //    let api_url = base_url + "/api/grf/dpbxcmd";
    //    let cmd: Cmd.CmdDropBox = { Action: Cmd.CmdDropBoxAction.DecryptTkn, Code: code }; 
    //    return HTTP.create_prms({ method: HTTP.HTTP_method.POST, restUrl: api_url, rawToken: t, data: JSON.stringify(cmd), contentType: HTTP.ContentType.application_json });
    //}
    on_save_item() {
        Api.save(this.state.cmd)
            .catch((err) => { this.set_info(Utl.any_to_str(err)); });
    }
    on_get_dp_token() {
        domains_1.Grf.getAccessTokenAsync()
            .then((t) => {
            this.on_get_dp_token_(t);
        })
            .catch((err) => { this.set_info(Utl.any_to_str(err)); });
    }
    on_get_dp_status() {
        const fn_name = "on_get_dp_status";
        Api.Dpbx.get_status()
            .then((r) => {
            let s = App.Cfg.get();
            if (!s) {
                // default 
                s = {
                    dpbx: { enabled: true, tkn: "" }
                };
                App.Cfg.set_save(s)
                    .catch((err) => {
                    App.Log.info(fn_name + Utl.any_to_str(err));
                });
            }
            //debugger;
            let is_log_in = s.dpbx.tkn ? true : false;
            let is_enabled = (s.dpbx.enabled && r.enabled);
            if (s.dpbx.tkn) {
                App.Log.info(fn_name + "Token ensrypted: " + s.dpbx.tkn);
                // decode 
                //debugger;
                Api.Dpbx.decrypt_tkn(s.dpbx.tkn)
                    .then((rd) => {
                    this.setState((ps) => (Object.assign({}, ps, { drop_box: Object.assign({}, ps.drop_box, { enabled: is_enabled, is_log_in: is_log_in, tkn: rd.tkn }) })));
                })
                    .catch((err) => {
                    this.setState((ps) => (Object.assign({}, ps, { drop_box: Object.assign({}, ps.drop_box, { enabled: is_enabled, is_log_in: false, tkn: "" }) })));
                    App.Log.info(fn_name + Utl.any_to_str(err));
                });
            }
            else {
                this.setState((ps) => (Object.assign({}, ps, { drop_box: Object.assign({}, ps.drop_box, { enabled: (s.dpbx.enabled && r.enabled), is_log_in: s.dpbx.tkn ? true : false }) })));
            }
        })
            .catch((err) => {
            this.setState((ps) => (Object.assign({}, ps, { drop_box: Object.assign({}, ps.drop_box, { enabled: false, is_log_in: false }) })));
            App.Log.info(fn_name + Utl.any_to_str(err));
        });
    }
    //on_set_dp_token() {
    //    Grf.getAccessTokenAsync()
    //        .then((t) => {
    //            this.drop_box_save_token(t, "");
    //        })
    //        .catch((err) => { this.set_info(Utl.any_to_str(err)); });
    //}
    on_get_dp_token_(t) {
        const fn_name = "on_get_dp_token";
        let url = this.get_url_ses();
        App.Log.info("URL:" + url);
        const base_url = url;
        let Api_Url = base_url + "/api/dp/tkn";
        HTTP.create_prms({ method: HTTP.HTTP_method.GET, restUrl: Api_Url, rawToken: "", contentType: HTTP.ContentType.application_json })
            .then((r) => {
            this.set_info("OK on_get_dp_token:" + Utl.any_to_str(r));
            let dlg = new DlgDropBox.Dialog();
            let purl = encodeURI(r.data);
            let dlg_url = "./dlgDropBox.html" + "?url_data=" + purl;
            let dlgp = Dlg.create_dlg_async(dlg_url, { height: 40, width: 50, displayInIframe: false, promptBeforeOpen: true }, dlg);
            dlgp
                .then((d) => {
                let is_log_in = false;
                let code = "";
                switch (d) {
                    case Dlg.Result.Cancel:
                        {
                        }
                    default:
                        {
                            let data = JSON.parse(d);
                            if (data.data) {
                                is_log_in = true;
                                code = data.data;
                            }
                            else if (data.error) {
                                App.Log.error(data.error, " err " + fn_name);
                            }
                            else {
                                App.Log.error(d, " uknown " + fn_name);
                            }
                        }
                }
                if (is_log_in && code.length != 0) {
                    Api.Dpbx.get_tkn_from_code(code)
                        .then((r) => {
                        this.setState((ps) => (Object.assign({}, ps, { drop_box: Object.assign({}, ps.drop_box, { is_log_in: true, tkn: r.tkn }) })));
                        App.Log.info(fn_name + "Token ensrypted: " + r.tkne);
                        var s = App.Cfg.get();
                        s.dpbx.tkn = r.tkne;
                        App.Cfg.set_save(s)
                            .catch((err) => {
                            App.Log.info(fn_name + Utl.any_to_str(err));
                        });
                    })
                        .catch((err) => {
                        App.Log.error(Utl.any_to_str(err), " Err get_tkn_from_code");
                    });
                    //this.drop_box_save_token(t, code)
                    //    .then(() => {
                    //        this.setState((ps) => ({ ...ps, drop_box: { ...ps.drop_box, is_log_in: true} }));
                    //    })
                    //    .catch((err) => {
                    //        App.Log.error(Utl.any_to_str(err), " Err drop_box_save_token");
                    //    })
                }
                dlg.close();
            })
                .catch((err) => {
                dlg.close();
                App.Log.error(Utl.any_to_str(err), " Err dlgDropBox");
            });
        })
            .catch((err) => {
            App.Log.error(err, " Err on_get_dp_token:");
            this.set_info(Utl.any_to_str(err) + "  Err on_get_dp_token:");
        });
    }
    render() {
        /*
        let msFilebr;
        if (this.state.cmd.SaveTo == Cmd.SaveTo.OneDtrive) {
            
            <GraphFileBrowser
                getAuthenticationToken={this.getAuthenticationToken}
                onSuccess={(selectedKeys: any[]) => console.log(selectedKeys)}
                onCancel={(err: Error) => console.log(err.message)}
            />

        }
        */
        let ext_to;
        if (this.state.drop_box.enabled) {
            ext_to = React.createElement("div", { style: { width: "170px" } },
                React.createElement(index_2.ComboBox, { defaultSelectedKey: models_1.Cmd.SaveTo[models_1.Cmd.SaveTo.OneDtrive], label: "Where to save", autoComplete: "off", options: INITIAL_OPTIONS_SaveTo, useComboBoxAsMenuWidth: true, onChange: (event, option, index, value) => {
                        let st = models_1.Cmd.SaveTo[option.key];
                        this.setState((ps) => (Object.assign({}, ps, { cmd: Object.assign({}, ps.cmd, { SaveTo: st }) })));
                        if (st == models_1.Cmd.SaveTo.DropBox && !this.state.drop_box.is_log_in) {
                            this.on_get_dp_token();
                        }
                    } }));
        }
        let ext;
        switch (this.state.cmd.SaveAs) {
            case models_1.Cmd.SaveAs.PDF:
                {
                    //ext = <div>PDF</div>
                }
                break;
            case models_1.Cmd.SaveAs.ZIP:
                {
                    ext = React.createElement("div", { style: { width: "170px" } },
                        React.createElement(index_2.ComboBox, { defaultSelectedKey: models_1.Cmd.RootItemConversion[models_1.Cmd.RootItemConversion.PDF], label: "Root Item Conversion", autoComplete: "off", options: INITIAL_OPTIONS_RootItemConversion, useComboBoxAsMenuWidth: true, onChange: (event, option, index, value) => {
                                this.setState((ps) => (Object.assign({}, ps, { cmd: Object.assign({}, ps.cmd, { RootItemConversion: models_1.Cmd.RootItemConversion[option.key] }) })));
                            } }));
                }
                break;
        }
        let ext2;
        if (this.state.attch.is_item_present) {
            switch (this.state.cmd.SaveAs) {
                //case Cmd.SaveAs.PDF:
                case models_1.Cmd.SaveAs.ZIP:
                    {
                        ext2 = React.createElement("div", { style: { width: "170px" } },
                            React.createElement(index_2.ComboBox, { defaultSelectedKey: models_1.Cmd.AttachedItemConversion[models_1.Cmd.AttachedItemConversion.MSG], label: "Attached Item Conversion", autoComplete: "off", options: INITIAL_OPTIONS_AttachedItemConversion, useComboBoxAsMenuWidth: true, onChange: (event, option, index, value) => {
                                    this.setState((ps) => (Object.assign({}, ps, { cmd: Object.assign({}, ps.cmd, { AttachedItemConversion: models_1.Cmd.AttachedItemConversion[option.key] }) })));
                                } }));
                    }
                    break;
            }
        }
        let ext3;
        if (this.state.attch.data.length > 0) {
            let data = this.state.attch.data;
            switch (this.state.cmd.SaveAs) {
                //case Cmd.SaveAs.PDF:
                case models_1.Cmd.SaveAs.ZIP:
                case models_1.Cmd.SaveAs.EML:
                case models_1.Cmd.SaveAs.MSG:
                    {
                        ext3 =
                            React.createElement("div", null,
                                React.createElement("div", null, "Attachments"),
                                React.createElement(MarqueeSelection_1.MarqueeSelection, { selection: this._selection_att },
                                    React.createElement(DetailsList_1.DetailsList, { items: data, columns: this._columns_att, setKey: "set", layoutMode: DetailsList_1.DetailsListLayoutMode.justified, selection: this._selection_att, selectionPreservedOnEmptyClick: true, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "Row checkbox" })));
                    }
                    break;
            }
        }
        return (React.createElement("div", null,
            React.createElement("div", null, this.state.info),
            ext_to,
            React.createElement("br", null),
            React.createElement("div", { style: { height: "10px" } }),
            React.createElement("div", { style: { width: "100px" } },
                React.createElement(index_2.ComboBox, { defaultSelectedKey: "MSG", label: "Save as", autoComplete: "off", options: INITIAL_OPTIONS_conversion_type, useComboBoxAsMenuWidth: true, onChange: (event, option, index, value) => {
                        this.setState((ps) => (Object.assign({}, ps, { cmd: Object.assign({}, ps.cmd, { SaveAs: models_1.Cmd.SaveAs[option.key], AttchmentProcessingKind: models_1.Cmd.SaveAs[option.key] == models_1.Cmd.SaveAs.PDF ? models_1.Cmd.AttchmentProcessingKind.None : ps.cmd.AttchmentProcessingKind }) })));
                    } })),
            React.createElement("br", null),
            React.createElement("div", { style: { height: "10px" } }),
            ext,
            React.createElement("div", { style: { height: "10px" } }),
            ext2,
            React.createElement("br", null),
            React.createElement("div", { style: { height: "10px" } }),
            ext3,
            React.createElement("div", { style: { height: "10px" } }),
            React.createElement(index_1.PrimaryButton
            //  style={{ float: "right" }}
            , { 
                //  style={{ float: "right" }}
                autoFocus: true, tabIndex: 1, onClick: () => { this.on_save_item(); }, text: "Doit" }),
            React.createElement("div", { style: { height: "10px" } })));
    }
}
const mapStateToProps = (state) => {
    return { addIn: state.addIn };
};
const mapDispatchToProps = (dispatch) => {
    return {
        actions: redux_1.bindActionCreators(utils_1.omit(actions_1.AddInActions, 'Type'), dispatch)
        //actions:
        //    {
        //        wait_on: bindActionCreators(SuspeActions.wait_on, dispatch),
        //        wait_off: bindActionCreators(SuspeActions.wait_off, dispatch)
        //    }
    };
};
exports.MainCmp = react_redux_1.connect(mapStateToProps, mapDispatchToProps)(MainCmp_);
