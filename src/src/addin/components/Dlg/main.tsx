import * as React from "react";
import * as ReactDOM from "react-dom";
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { MainCmp } from './MainCmp';

import * as Dlg from "../../../domains/dlg";

const app_root_id = "app_dlg_dlg"




let fn_init = () => {



    Office.initialize = (reason) => {

        let start = (f: () => void) => {
            if (document.readyState != "complete") {
                window.addEventListener("load", f);
            }
            else {
                f();
            }
        };

        start(function () {


                ReactDOM.render(
                    <Fabric>
                    <MainCmp />
                    </Fabric>
                    ,
                    document.getElementById(app_root_id)
                );

            

        });

    }
}

fn_init();





