import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest, scharenIn): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    let message = "";
    message += `Scharen: ${scharenIn.length}\n`

    let count_tn = 0;
    let count_leader = 0;

    for (let schar of scharenIn) {
        count_tn += parseInt(schar.num_tn);
        count_leader += parseInt(schar.num_leader);
    }

    message += `TN: ${count_tn}\n`
    message += `Leiter: ${count_leader}\n`
    message += `Total: ${count_tn + count_leader}\n`

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: message
    };

};

export default httpTrigger;