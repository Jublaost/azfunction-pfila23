import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getToken } from "../Common/Token";
import axios, { AxiosRequestConfig } from 'axios';
import { validateRECAP } from "../Common/Recaptcha";

const SITE_ID = process.env.SITE_ID;
const LIST_ID = process.env.HELFENDEFEST_LIST_ID;

const MS_GRAPH_ENDPOINT_LISTITEM = 'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/lists/' + LIST_ID + '/items';
const MS_GRAPH_ENDPOINT_SENDMAIL = 'https://graph.microsoft.com/v1.0/users/ok@pfila23.ch/sendMail';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    let token = await getToken();
    //context.log("Token: ", token)
    context.log("Body: ", req.body)

    const unixTime = Math.floor(Date.now() / 1000)

    let validation = await validateRECAP(context, req.body["g-recaptcha-response"]);

    if (!validation) {
        context.log("validation failed");
        context.res = {
            status: 500,
            message: "Recaptcha failed"
        }
        return
    } else {
        context.log("Recaptcha Succesful validated");
    }

    try {
        let response = await postListItem(token, req.body);
        context.log("Status: ", response.status);
        if (response.status == 201 || response.status == 200) {
            context.log(response.data);
            let mail = await sendMail(token, req.body);
            context.log("Mail send", mail)
            context.res = {
                status: 200,
            };
        } else {
            context.log("Error: ", response)
            context.res = {
                status: 500,
                body: "server error"
            };
        }
    } catch (e) {
        context.log(e)
        context.res = {
            status: 500,
            body: "server error"
        };
    }
};

export default httpTrigger;


async function postListItem(token: string, body: any): Promise<any> {
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT_LISTITEM,
        headers: {
            'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: {
            "fields": {
                "Title": body.vorname + ' ' + body.nachname,
                "Geburtstag": body.geburtstag
            }
        }
    }

    return await axios(config)
        .then(response => {
            return response;
        })
        .catch(error => {
            return error;
        });
}


async function sendMail(token: string, body: any) {
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT_SENDMAIL,
        headers: {
            'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: {
            "message": {
                "subject": "Bestätigung Anmeldung Pfila23",
                "body": {
                    "contentType": "html",
                    "content": 
                    "Liebe*r " + body.vorname + "<br /><br /><p><strong>Ganz herzlichen Dank für deine grossartige Hilfe beim Pfila 23!</strong></p><p>Wir freuen uns riesig, dich auf unserem Helferfest begrüssen zu dürfen. Dein Engagement und deine Unterstützung waren unerlässlich, um dieses tolle Erlebnis möglich zu machen. Gemeinsam wollen wir nun einen stimmungsvollen Abschluss feiern.</p><p>Falls du allgemeine Fragen hast, zögere bitte nicht, uns zu kontaktieren: <a href='mailto:ok@pfila23.ch'>ok@pfila23.ch</a></p><p>Wir freuen uns schon auf dich und schicken dir bis dahin jublastische Grüsse</p><p>Tanurana und Tilion<br />Hauptleitung</p>"
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": body.email
                        }
                    }
                ]
            },
            "saveToSentItems": "true"
        }
    }

    return await axios(config)
        .then(response => {
            return response;
        })
        .catch(error => {
            return error;
        });
}