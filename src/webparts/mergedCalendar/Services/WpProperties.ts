import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPPermission } from "@microsoft/sp-page-context";
import {SPHttpClientResponse, SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

const getWpSavedData = (context: WebPartContext) : Promise <{}[]> => {
    const apiUrl = context.pageContext.web.absoluteUrl + "/_api/sitepages/pages("+ context.pageContext.listItem.id +")";

    return new Promise <{}[]> (async(resolve, reject) =>{
        context.spHttpClient
        .get(apiUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse)=>{
            response.json().then((results:any)=>{
                resolve(results);
            });
        });
    });
};

const updateWpData = async (context: WebPartContext, propName: string, propVal: any) : Promise <any> => {
    const apiUrl = context.pageContext.web.absoluteUrl + "/_api/sitepages/pages("+ context.pageContext.listItem.id +")";
    const _savedData = await context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1); //get saved webpart data
    let newCanvasContent: any = null;

    if(_savedData.ok){
        const _results = await _savedData.json();
        if(_results){
            const canvasContent : any = JSON.parse(_results.CanvasContent1);
            for (const v of canvasContent){
                if (v.id === context.instanceId){
                    newCanvasContent = canvasContent;
                    newCanvasContent[0].webPartData.properties[propName] = propVal;
                    break;
                }
            }
        }
    }    
    return JSON.stringify(newCanvasContent);
};

export const setWpData = async (context: WebPartContext, propName: string, propVal: any) : Promise <any> =>{
    //get updated data
    let _updatedWpData = await updateWpData(context, propName, propVal);
    const spOptions : ISPHttpClientOptions = {
        body: `{"__metadata":{"type":"SP.Publishing.SitePage"},"CanvasContent1": ${JSON.stringify(_updatedWpData)}}`
    };

    //checkout the page first
    let apiUrl = `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${context.pageContext.listItem.id})/checkoutpage`;
    const _pageCheckout = await context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {});

    //save page with new data
    apiUrl = `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${context.pageContext.listItem.id})/savepage`;
    const _savedData = await context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOptions);
    if (_savedData.ok){
        console.log("Page Webpart data is updated!");
    }else{
        console.log(`Error: ${_savedData.statusText}`);
    }

    //publishing the page 
    apiUrl = `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${context.pageContext.listItem.id})/publish`;
    const _checkedInPage = await context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {});
};


export const isUserManage = (context: WebPartContext) : boolean =>{
    const userPermissions = context.pageContext.web.permissions,
        permission = new SPPermission (userPermissions.value);
    
    return permission.hasPermission(SPPermission.manageWeb);
};