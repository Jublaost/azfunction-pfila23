import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest, scharenIn): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    context.log('Scharen:', scharenIn)

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: scharenIn
    };

};

export default httpTrigger;