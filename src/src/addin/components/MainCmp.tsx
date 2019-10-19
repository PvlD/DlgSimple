import * as React from 'react';
import * as DlgSimple from "../components/dlg/dlg_wrp"
import * as Dlg from "../../domains/dlg"

import {
    PrimaryButton,
} from 'office-ui-fabric-react/lib/index';




export class MainCmp extends React.Component<{}, any> {

    constructor(props: {}, context?: any) {
        super(props, context);

    }

    show_dlg() {
        


        let dlg = new DlgSimple.Dialog();
        let dlg_url = "./DlgSimple.html";

        let dlgp = Dlg.create_dlg_async(dlg_url, { height: 40, width: 50, displayInIframe: false, promptBeforeOpen: true }, dlg);
        dlgp
            .then((d) => {
               

                let dd= JSON.parse(d);
                switch (dd) {
                    case Dlg.Result.Cancel:
                        {
                            console.info("Cancel");
                        }
                        break;
                    case Dlg.Result.Yes:
                        {
                            console.info("Yes");
                        }
                        break;
                    case Dlg.Result.No:
                        {
                            console.info("No");
                        }
                        break;
                    default:
                        {
                            console.error("Unknown Dlg.Result");
                        }
                }


                dlg.close();

            })
            .catch((err) => {
                dlg.close();
                console.error(err);
            })

    }

    render() {


        return (

            

            <div>
                <div style={{ height: "10px" }} />
                <PrimaryButton
                    
                    autoFocus={true}
                    tabIndex={1}
                    onClick={() => { this.show_dlg(); }}
                    text="Show Dialog"
                />

            </div>

        );
    }
}

