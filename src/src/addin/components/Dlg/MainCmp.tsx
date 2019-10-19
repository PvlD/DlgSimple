import * as React from 'react';

//import * as  DlgDropBox from "../components/DlgDropBox/dlg_wrp";
//import * as Dlg from "../../domains/dlg";

import {
    PrimaryButton,
} from 'office-ui-fabric-react/lib/index';
import { Dlg } from '../../../domains';









export class MainCmp extends React.Component<{}, {}> {

    constructor(props: {}, context?: any) {
        super(props, context);

    }

    render() {


        return (

            <div>
                <div style={{ marginLeft: "10px" }}>
                <div style={{ height: "10px" }} />
                <PrimaryButton
                    
                    autoFocus={true}
                    tabIndex={1}
                    onClick={() => { Office.context.ui.messageParent(JSON.stringify(Dlg.Result.Yes)); }}
                    text="Yes"
                />
                <div style={{ height: "10px" }} />
                <PrimaryButton

                    autoFocus={true}
                    tabIndex={1}
                    onClick={() => { Office.context.ui.messageParent(JSON.stringify(Dlg.Result.No)); }}
                    text="No"
                    />
                </div>
            </div>
            
        );
    }
}

