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

        case ("Aluminium"):
            colorHex = "#";
            break;
        case ("Aniseed"):
            colorHex = "#c3d329";
            break;
        case ("Autumn"):
            colorHex = "#a4905e";
            break;
        case ("Black"):
            colorHex = "#302a26";
            break;
        case ("Boulder"):
            colorHex = "#b8bdbe";
            break;
        case ("Burgundy"):
            colorHex = "#5f2d37";
            break;
        case ("Buttercup"):
            colorHex = "#fad825";
            break;
        case ("Camel"):
            colorHex = "#cab08e";
            break;
        case ("Carrot"):
            colorHex = "#e46138";
            break;
        case ("Celadon"):
            colorHex = "#90cce1";
            break;
        case ("Celestial Blue"):
            colorHex = "#036184";
            break;
        case ("Champagne"):
            colorHex = "#f6eac3";
            break;
        case ("Cocoa"):
            colorHex = "#675d4b";
            break;
        case ("Concrete"):
            colorHex = "#69666f";
            break;
        case ("Dark Blue"):
            colorHex = "#225b87";
            break;
        case ("Dijon"):
            colorHex = "#e2a135";
            break;
        case ("Hemp"):
            colorHex = "#e9dbb5";
            break;
        case ("Lagoon"):
            colorHex = "#4097b6";
            break;
        case ("Lemon"):
            colorHex = "#f1e378";
            break;
        case ("Marine"):
            colorHex = "#393e5d";
            break;
        case ("Midnight Blue"):
            colorHex = "#384e87";
            break;
        case ("Moss Green"):
            colorHex = "#829975";
            break;
        case ("Olive"):
            colorHex = "#727261";
            break;
        case ("Pepper"):
            colorHex = "#b19b79";
            break;
        case ("Poppy"):
            colorHex = "#bb272d";
            break;
        case ("Porcelain Green"):
            colorHex = "#169096";
            break;
        case ("Raspberry"):
            colorHex = "#c04255";
            break;
        case ("Sandy Beige"):
            colorHex = "#c9b7ae";
            break;
        case ("Spruce"):
            colorHex = "#264c43";
            break;
        case ("Steel Blue"):
            colorHex = "#5ba1b7";
            break;
        case ("Teak"):
            colorHex = "#6a635c";
            break;
        case ("Tennis Green"):
            colorHex = "#17594a";
            break;
        case ("Terracotta"):
            colorHex = "#894e43";
            break;
        case ("Thistle Blue"):
            colorHex = "#0a5f7b";
            break;
        case ("Vanilla"):
            colorHex = "#e6d2af";
            break;
        case ("Velvet Red"):
            colorHex = "#954339";
            break;
        case ("Victoria Blue"):
            colorHex = "#0061a3";
            break;
        case ("Walnut Stain"):
            colorHex = "#4b3429";
            break;
        case ("White"):
            colorHex = "#f2f4f8";
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
                            BgColor: result.BgColor.replace(/ /g,''),
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
                            Disabled: (result.CalType == 'My School' || result.CalendarType == 'My School') ? true : false,
                            Dpd: (result.CalType == 'Rotary' || result.CalendarType == 'Rotary') ? true : false,
                            //LegendURL : result.CalType !== 'Graph' ? result.CalURL + "/Lists/" + result.CalName : null
                            LegendURL : result.Link || "",
                            View: result.View || ''
                        });
                    });                    
                    resolve(calSettings);
                });
                
            });
    });
};

export const updateCalSettingsX = (context:WebPartContext, listName: string, calSettings:any, checked?:boolean, dpdCalName?:any) : Promise <any> =>{
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

export const updateCalSettings = (context:WebPartContext, listName: string, calSettings:any, checked?:boolean, dpdCalName?:any, rotaryCals?:any) : Promise <any> =>{
    // console.log("checked", checked);
    // console.log("dpdCalName", dpdCalName);
    // console.log("rotaryCals", rotaryCals);

    const rotaryCal = rotaryCals ? rotaryCals.filter(item => item.Title === dpdCalName)[0] : null;

    let restApiUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('"+listName+"')/items("+calSettings.Id+")",
        body: string = JSON.stringify({
            Title: calSettings.Title,
            ShowCal: checked,
            CalName: dpdCalName ? dpdCalName : calSettings.CalName,
            CalURL: dpdCalName ? rotaryCal.CalURL : calSettings.CalURL
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

export const isUserGrpCal = (calView: string, posGrps: any, userGrps: any) => {
    if (posGrps[calView.trim()] == undefined) return true;
    else{
        for (let userGrp of userGrps){
            if (posGrps[calView.trim()] && posGrps[calView.trim()].indexOf(Number(userGrp)) !== -1){
                return true;
            }
        }
        return false;
    }
};
export const isPosGrpsCal = (calSettings) => {
    return calSettings.filter(item => item.View !== 'AllItems' && item.View !== '' && item.Chkd).length > 0;
};

