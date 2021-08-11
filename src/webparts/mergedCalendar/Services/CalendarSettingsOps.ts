import { WebPartContext } from "@microsoft/sp-webpart-base";
import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from "@microsoft/sp-http";


const getColorHex = (colorName:string) : string => {
    let colorHex : string;
    switch (colorName) {
        case ("Black"):
            colorHex = "#000000";
            break;
        case ("Blue"):
            colorHex = "#0096CF";
            break;
        case ("Green"):
            colorHex = "#27AE60";
            break;
        case ("Grey"):
            colorHex = "#9FA7A7";
            break;
        case ("Mint"):
            colorHex = "#1C9A82";
            break;
        case ("Navy"):
            colorHex = "#4C5F79";
            break;
        case ("Orange"):
            colorHex = "#EA8020";
            break;
        case ("Pink"):
            colorHex = "#F46C9E";
            break;
        case ("Purple"):
            colorHex = "#A061BA";
            break;
        case ("Red"):
            colorHex = "#D7574A";
            break;
        case ("Teal"):
            colorHex = "#38A8AC";
            break;
        case ("White"):
            colorHex = "#FFFFFF";
            break;
        case ("Yellow"):
            colorHex = "#DAA62F";
            break;
    }
    return colorHex;
};

export const getCalSettings = (context:WebPartContext, listName: string) : Promise <{}[]> => {
    
    console.log('Get Cal Settings Function');

    let restApiUrl : string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('"+listName+"')/items" ,
        calSettings : {}[] = [];

    return new Promise <{}[]> (async(resolve, reject)=>{
        context.spHttpClient
            .get(restApiUrl, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse)=>{
                response.json().then((results:any)=>{
                    results.value.map((result:any)=>{
                        calSettings.push({
                            BgColor: result.BgColor,
                            BgColorHex : getColorHex(result.BgColor),
                            CalName : result.CalName,
                            CalType: result.CalendarType ? result.CalendarType : result.CalType,
                            CalURL: result.CalURL,
                            FgColor: result.FgColor,
                            FgColorHex: getColorHex(result.FgColor),
                            Id: result.Id,
                            ShowCal: result.ShowCal,
                            Title: result.Title,
                            Chkd: result.ShowCal ? true : false,
                            Disabled: result.CalType == 'My School' ? true : false,
                            Dpd: result.CalType == 'Rotary' ? true : false,
                            //LegendURL : result.CalType !== 'Graph' ? result.CalURL + "/Lists/" + result.CalName : null
                            LegendURL : result.Link || ""
                        });
                    });                    
                    resolve(calSettings);
                });
                
            });
    });
};

export const updateCalSettings = (context:WebPartContext, listName: string, calSettings:any, checked?:boolean, dpdCalName?:any) : Promise <any> =>{
    let restApiUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('"+listName+"')/items("+calSettings.Id+")",
        body: string = JSON.stringify({
            Title: calSettings.Title,
            ShowCal: checked,
            CalName: dpdCalName ? dpdCalName : calSettings.CalName
        }),
        options: ISPHttpClientOptions = {
            headers:{
                Accept: "application/json;odata=nometadata", 
                "Content-Type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",                
            },
            body: body
        };

    return new Promise <string> (async(resolve, reject)=>{
        context.spHttpClient
        .post(restApiUrl, SPHttpClient.configurations.v1, options)
        .then((response: SPHttpClientResponse)=>{
            //console.log('item updated !!');
            resolve("Item updated");
        });
    });
};


