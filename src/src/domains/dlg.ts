
import { On_Complete_CallBack, excep_wrap} from "./utl"

export enum Result {
    Yes = "Yes",
    No = "No",
    Cancel = "Cancel"
}


export interface Error { error: string; };
export interface Data { data: string; };

export type ResultDlg = Result | string;



export interface I_Dlg_messageHandler_arg {
    message: string;
}

export interface I_Dlg_eventHandler_arg {
    error: number;
}
type Resolve<TRslt> = (value?: TRslt | OfficeExtension.IPromise<TRslt>) => void;
type Reject = (error?: any) => void;


export abstract class OffcDlg_wrp<TRslt>
{
    m_dialog: Office.Dialog;
    m_on_resolve: Resolve<TRslt>;
    m_on_reject: Reject;

    constructor() {
    }

    set dialog(dialog: Office.Dialog) {
        this.m_dialog = dialog;
        this.m_dialog.addEventHandler(Office.EventType.DialogMessageReceived, (a: any) => { this.on_message(a as I_Dlg_messageHandler_arg); })
        this.m_dialog.addEventHandler(Office.EventType.DialogEventReceived, (a: any) => { this.on_event(a as I_Dlg_eventHandler_arg); })

    }
    get dialog() { return this.m_dialog }; 

    set on_resolve(h: Resolve<TRslt>) { this.m_on_resolve = h; }
    get on_resolve() { return this.m_on_resolve; }

    set on_reject(h: Reject) { this.m_on_reject = h; }
    get on_reject() { return this.m_on_reject; }


    abstract on_message(a: I_Dlg_messageHandler_arg): void;
    abstract on_event(a: I_Dlg_eventHandler_arg): void;

    close() {
        this.m_dialog.close(); 
    }
}

export async function create_dlg_async<TRslt>(dlg_url_relative: string, options: Office.DialogOptions, dlg: OffcDlg_wrp<TRslt>): Promise<TRslt> {

    let loc = window.location.href.lastIndexOf("/");
    let dlg_url = window.location.href.substring(0, loc) + "/" + dlg_url_relative;

    return create_dlg_async_abs(dlg_url, options, dlg);

}

export async function create_dlg_async_abs<TRslt>(dlg_url: string, options: Office.DialogOptions, dlg: OffcDlg_wrp<TRslt>): Promise<TRslt> {

    let fn = function (cllbck: On_Complete_CallBack) {
        Office.context.ui.displayDialogAsync(dlg_url, options, cllbck);
    }

    let pr = new Promise<TRslt>((resolve, reject) => {
        fn((r: Office.AsyncResult<any>) => {
            let f = () => {
                if (r.status == Office.AsyncResultStatus.Failed) {
                    reject(r.error);
                    return;
                }
                dlg.on_resolve = resolve;
                dlg.on_reject = reject;
                dlg.dialog = r.value as Office.Dialog;
            }
            excep_wrap(f, (ex) => { reject(ex); })
        });
    });
    return pr;

}


export namespace Dlg_YNC {

    export function on_yes() {
        Office.context.ui.messageParent(Result.Yes);
    }

    
    export  function on_no() {
        Office.context.ui.messageParent(Result.No);
    }

    export  function on_cancel() {
        Office.context.ui.messageParent(Result.Cancel);
    }

    export function on_data(d: string) {
        Office.context.ui.messageParent(d);
    }




    export class Dialog extends OffcDlg_wrp<ResultDlg>
    {
        name = "";

        constructor(name: string) {
            super();
            this.name = name;
        }


        on_message(a: I_Dlg_messageHandler_arg): void {
            switch (a.message) {
                case Result.Yes:
                    this.on_resolve(Result.Yes);
                    
                    break;
                case Result.No:
                    this.on_resolve(Result.No);
                    
                    break;
                default:
                    this.on_resolve(a.message);
                    break;
            }

        }
        on_event(a: I_Dlg_eventHandler_arg): void {

            switch (a.error) {
                case 12002:
                case 12003:
                    this.on_reject(name + " error:" + a.error.toString());
                    break;
                case 12006:
                    // The dialog was closed, typically because the user the pressed X button.
                    this.on_resolve(Result.Cancel);
                    break;
                default:
                    this.on_reject(name + " Undefined error:" + a.error.toString());
                    break;
            }

        }
    }
}

