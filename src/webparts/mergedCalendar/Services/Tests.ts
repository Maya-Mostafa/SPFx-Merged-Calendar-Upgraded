import { WebPartContext } from "@microsoft/sp-webpart-base";
import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, AadHttpClient, HttpClientResponse, HttpClient, IHttpClientOptions, MSGraphClient} from "@microsoft/sp-http";

export class Tests{

    private apiClient : AadHttpClient;
    public getExt(context:WebPartContext):any{
         const appClientID :string = "3b7def80-9920-47b6-924d-1edcab90a211";
         let apiUrl :string = "https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items";

        // const appClientID :string = "eb994916-2c73-4bc6-b4bd-c945f62eac26";
        // let apiUrl :string = "https://pdsbserviceapi.azurewebsites.net/api/wcf/GetLunchRoomSupByLocation?LocationId=1415";

        //const appClientID :string = "2eb4dc90-8ac5-414b-ab9f-0be1be2c0b61";
        //let apiUrl :string = "https://graph.microsoft.com/v1.0/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location";
        
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
            context.aadHttpClientFactory
                .getClient(appClientID)
                .then((client: AadHttpClient)=>{
                    this.apiClient = client;
                    console.log(client);
                    resolve();

                    this.apiClient
                        .get(apiUrl, AadHttpClient.configurations.v1)
                        .then((res:HttpClientResponse) : Promise<any>=>{
                            return res.json().then((results:any)=>{
                                console.log(results);
                            });
                        });

                },err => reject(err));
        });
    }
    public getExtSchool(context:WebPartContext){

        //working :)
        /*const options = {
            headers : { 
                'Accept': 'application/json;odata=verbose'
            }
        }
        fetch("https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items", options).then((response)=>{
            response.json().then((data)=>{
                console.log(data)
            })
            
        })*/
        
        /*context.aadHttpClientFactory
            .getClient("https://pdsb1.azure-api.net")
            .then((client: AadHttpClient):void =>{
                this.apiClient = client;
                console.log(client)
                this.apiClient
                    .get("https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items", AadHttpClient.configurations.v1)
                    .then((res:HttpClientResponse) : Promise <any>=>{
                        return res.json().then((results:any)=>{
                            console.log(results);
                        })
                    })
        })*/

        //https://schools.peelschools.org/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items
        //https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items

        //working :)
        /*const options:IHttpClientOptions = {
            headers : { 
                'Accept': 'application/json;odata=verbose'
            }
        }
        context.httpClient
            .get("https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items", HttpClient.configurations.v1, options)
            .then((res: HttpClientResponse): Promise<any> => {
                return res.json().then((results:any)=>{
                    console.log(results)
                });
        })*/

        /*const requestHeaders: Headers = new Headers();
        requestHeaders.append('Accept', 'application/json;odata=verbose');
        const options:ISPHttpClientOptions = {
            headers : requestHeaders,
        }
        return new Promise<any>(async(resolve, reject)=>{
            context.spHttpClient
                .get("https://pdsb1.azure-api.net/peelschools/sec/johnfraser/_api/web/lists/getByTitle('Calendar')/items", SPHttpClient.configurations.v1, options)
                .then((response: SPHttpClientResponse)=>{
                    response.json().then((results:any)=>{
                        console.log(results);
                        resolve(results);
                    })
                })
        })*/

    }

}