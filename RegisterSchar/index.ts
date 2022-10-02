import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const { v4: uuidv4 } = require('uuid')

const RECAPTCHA = process.env["recaptchaCodev3"]
const APP_ID = process.env["APP_ID"];
const APP_SECRET = process.env["APP_SECRET"];
const TENANT_ID = process.env["TENANT_ID"];

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT_SENDMAIL = 'https://graph.microsoft.com/v1.0/users/ok@pfila23.ch/sendMail';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    context.log("Body: ", req.body);

    let validation = await validateRECAP(context, req.body["g-recaptcha-response"]);

    if (!validation) {
        context.log("validation failed");
        context.res = {
            status: 500
        }
        return
    }

    let uuid = uuidv4();
    let schar = req.body;
    schar.id = uuid;

    context.log("Schar: ", schar);

    try {
        context.bindings.scharOut = schar;

        let token = await getToken();
        context.log("Token: ", token);

        let mail = await sendMail(token, schar);
        context.log("Mail: ", mail);

        context.res = {
            status: 200,
            body: "successful"
        }
        return
    } catch (e) {
        context.res = {
            status: 500,
            body: "server error"
        }
        return
    }

};

export default httpTrigger;

async function validateRECAP(context: Context, token: string) {
    let config: AxiosRequestConfig = {
        method: 'post',
        url: "https://www.google.com/recaptcha/api/siteverify",
        params: {
            secret: RECAPTCHA,
            response: token
        }
    }
    return await axios(config)
        .then(response => {
            return response.data.success;
        })
        .catch(error => {
            context.log(error);
        });
}


/**
 * Get Token for MS Graph
 */
async function getToken(): Promise<string> {
    const postData = {
        client_id: APP_ID,
        scope: MS_GRAPH_SCOPE,
        client_secret: APP_SECRET,
        grant_type: 'client_credentials'
    };

    return await axios
        .post(TOKEN_ENDPOINT, qs.stringify(postData))
        .then(response => {
            return response.data.access_token;
        })
        .catch(error => {
            console.log(error);
        });
}

/**
 * Send Verification Email
 * @param token MS Graph Token
 * @param joinedUser joinedUser Object
 * @returns 
 */
async function sendMail(token: string, schar: any) {
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT_SENDMAIL,
        headers: {
            'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: {
            "message": {
                "subject": "Schön seid ihr dabei!",
                "body": {
                    "contentType": "html",
                    "content": "Hallo!<br /><br />Cool hast du deine Schar '" + schar.name + "' angemeldet!<br />Bei Fragen oder unklarheiten kannst du auf diese Mail antworten oder direkt: <a href='mailto:ok@pfila23.ch'>Pfila23 OK</a><br /><br />Jublastische Grüsse<br />Das Pfila23 Team"
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": schar.mail
                        }
                    }
                ]
            },
            "saveToSentItems": "true"
        }
    }

    return await axios(config)
        .then(response => {
            return response.data.value;
        })
        .catch(error => {
            console.log(error);
        });
}