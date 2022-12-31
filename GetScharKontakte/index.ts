import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest, scharenIn): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    let message = "";
    let emailliste = "";

    for (let schar of scharenIn) {
        message += schar.schar + '\n';
        message += schar.kontakt + '\n';
        message += schar.email + '\n\n';
        emailliste += schar.email + ";"
    }

    message += "\n\n\n\n" + emailliste;

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: message
    };

};

export default httpTrigger;