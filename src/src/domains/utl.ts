



const attemp_num_max = 1;
const attemp_timeout = 2000;
export function delay(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
export async function retry_promise(cpf: () => Promise<any>) {
    let error;

    for (let i = 0; i < attemp_num_max; i++) {

        let p = cpf();
        try {
            if (i == 0) {
                return await p;
            }
            await delay(attemp_timeout);
            return await p;
        } catch (err) {
            error = err;
        }
    }

    throw error;
}







export const update = (target: any, source: any) => {
    for (var attr in source) {
        if (target.hasOwnProperty(attr)) target[attr] = source[attr];
    }
};




export function xml_decode(s: string) {
    return s.replace(/&apos;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&gt;/g, '>')
        .replace(/&lt;/g, '<')
        .replace(/&amp;/g, '&')
        .replace(/&#xD;/g, '/r')
}

export function xml_encode(s: string) {
    return s.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    //.replace(/\r/g, '&#xD;');
}


export function excep_wrap(f: () => void, on_ex: (ex: any) => void) {
    try {
        f();
    } catch (ex) {
        on_ex(ex);
    }
}


export function err_as_string(r: Office.AsyncResult<any>) {
    let s = "";
    if (r.status == Office.AsyncResultStatus.Failed) {
        s += (" " + r.status.toString());
        s += (" error name:" + r.error.name);
        s += ("  message:" + r.error.message);
        s += ("  code:" + r.error.code);
    }
    return s;
}

export function any_to_str(o: any) {
    if (o== null) { return "null" }
    if (typeof o === "string") {
        return o;
    }
    if (typeof o === 'object') {
        return JSON.stringify(o)
    }
    return (o as Object).toString();
}

export function make_err(name : string, err: string | any) {

    let r: Error = { message: "", name:"" };
    if (name) {
        r.name = name;
    }
    r.message = any_to_str(err);
    return r;
}



export type On_Complete_CallBack = (r: Office.AsyncResult<any>) => void

//export function create_prms<TRslt>(fn: (cllbck: On_Complete_CallBack) => void): Promise<TRslt> {
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


export interface IDataError<TD>
{
    data?: TD,
    error?:any
}

export function wrap_data_error<TRslt>(p: Promise<TRslt>) {
    return p.then((d) => ({ data: d } as IDataError<TRslt> ))
        .catch((err) => ({ error: err } as IDataError<TRslt> )) 
}

export function un_wrap_data<TRslt>(p: any) {
    if (!p.data) throw  "can't un_wrap_data"
    return p.data as TRslt;

}



export function create_prms<TRslt>(fn: (cllbck: On_Complete_CallBack) => void): Promise<TRslt> {
    return create_prms_3<TRslt>(fn, (v: any)=>{ return v as TRslt });
}

 export function create_prms_3<TRslt>(fn: (cllbck: On_Complete_CallBack) => void, rslt_value_convrter: (v: any) => TRslt ): Promise<TRslt> {
    let pr = new Promise<TRslt>((resolve, reject) => {
        fn((r: Office.AsyncResult<any>) => {
            let f = () => {
                if (r.status == Office.AsyncResultStatus.Failed) {
                    reject(r.error);
                    return;
                }
                resolve(rslt_value_convrter(r.value));
            }
            excep_wrap(f, (ex) => { reject(ex); })
        });
    });
    return pr;
}

export function addHandlerAsync(mlbx: Office.Mailbox, eventType: Office.EventType, handler: (type: Office.EventType) => void) {
    return create_prms_3<boolean>((cbk: On_Complete_CallBack) => { mlbx.addHandlerAsync(eventType, handler, {}, cbk) }, (v: any) => { return true; });
}




//export type Dlg_messageHandler<TRslt> = (arg: I_Dlg_messageHandler_arg, func: (resolve: (value?: TRslt | OfficeExtension.IPromise<TRslt>) => void, reject: (error?: any) => void) => void) => void;
//export type Dlg_eventHandler<TRslt> = (arg: I_Dlg_eventHandler_arg, func: (resolve: (value?: TRslt | OfficeExtension.IPromise<TRslt>) => void, reject: (error?: any) => void) => void) => void;





export function error_msg(a: { error: any }[]) {
    if (!a) return null;
    let msg = '';
    a.forEach((v) => {
        if (v.error) {
            msg += any_to_str(v.error);
            msg += " ";
        }
    })
    if (msg == '') return null;
    return msg;
}


export function  IsNullOrEmpty(s: string)
{
    if (s == null) return true;
    return s.length == 0;
}

export function IsNullOrEmptyArr(s:any[] ) {
    if (s == null) return true;
    return s.length == 0;
}




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

