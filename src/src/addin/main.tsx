import * as React from "react";
import * as ReactDOM from "react-dom";
import { MainCmp } from './components/MainCmp';


const app_root_id = "dlg_app"



let fn_init = () => {

    

    Office.onReady().then((reason) => {


        let ff = function () {

            ReactDOM.render(
                 <MainCmp />
                ,
                document.getElementById(app_root_id)
            );

        };

        ff();

    })
}

fn_init();



    

    
    
    
    
    
    
    
