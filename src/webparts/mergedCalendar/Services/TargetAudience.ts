import { WebPartContext } from '@microsoft/sp-webpart-base';
import {MSGraphClient} from "@microsoft/sp-http";

export const getGraphMemberOf = async (context: WebPartContext) : Promise <string> =>{
    const graphUrl = '/me/transitiveMemberOf/microsoft.graph.group';
    //let graphUrl = '/me/memberof';

    return new Promise <string> ((resolve, reject)=>{
        context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient)=>{
            client
                .api(graphUrl)
                .header('ConsistencyLevel', 'eventual')
                .count(true)
                // .select('displayName')
                .top(500)
                .get((error, response: any, rawResponse?: any)=>{
                    //console.log("graph response", response);
                    resolve(response.value);
                });
        });
    });
};

export const isMember = async (context: WebPartContext) => {
    const userId = context.pageContext.user;
    const responseUrl = context.pageContext.web.absoluteUrl + "/_api/web/sitegroups/getByName('" + "Learning Technology Support Services - DL" + "')/Users";
};

export const isFromTargetAudience = (graphResponse: any, wpTargetAudience: any) => {
    const userGroups = [];
    for (const group of graphResponse){
        userGroups[group.displayName] = group.displayName;
    }
    for (const audience of wpTargetAudience){
        if (userGroups[audience.fullName])
            return true;
    }
    return false;
};

// Learning Technology Support Services - DL