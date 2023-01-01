import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getToken } from "../Common/Token";
import axios, { AxiosRequestConfig } from 'axios';

const SITE_ID = process.env.SITE_ID;
const SITE_ASSETS_ID = process.env.SITE_ASSETS_ID;
const LIST_ID = process.env.TN_LIST_ID;

const MS_GRAPH_ENDPOINT_LISTITEM = 'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/lists/' + LIST_ID + '/items';
const MS_GRAPH_ENDPOINT_UPLOAD = 'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/drives/' + SITE_ASSETS_ID + '/root:/Lists/' + LIST_ID + '/';
const MS_GRAPH_ENDPOINT_SENDMAIL = 'https://graph.microsoft.com/v1.0/users/ok@pfila23.ch/sendMail';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    let token = await getToken();
    //context.log("Token: ", token)
    context.log(req.body)

    try {
        if (req.body.signature) {
            let signatureFileUpload = await uploadFile(token, req.body.signature, req.body.vorname + '-' + req.body.nachname + '-signature')
            context.log("Signatur Upload: ", signatureFileUpload);
        }
        if (req.body.impfausweis) {
            let impfausweisFileUpload = await uploadFile(token, req.body.impfausweis, req.body.vorname + '-' + req.body.nachname + '-impfausweis')
            context.log("Impfausweis Upload: ", impfausweisFileUpload);
        }

        let response = await postListItem(token, req.body);
        context.log("Item: ", response)
        let mail = await sendMail(token, req.body);
        context.log("Mail send", mail)

        context.res = {
            status: 200, /* Defaults to 200 */
            body: req.body
        };
    } catch (e) {
        context.log(e)
        context.res = {
            status: 500,
            body: "server error"
        };
    }
};

export default httpTrigger;

async function uploadFile(token: string, file: string, filename: string) {
    let filePrefix = file.split(';')[0];
    let fileContentType = filePrefix.split(':')[1];
    let fileExtenstion = fileContentType.split('/')[1];

    const byteArray = Buffer.from(file.replace(/^[\w\d;:\/]+base64\,/g, ''), 'base64');

    let config: AxiosRequestConfig = {
        method: 'put',
        url: MS_GRAPH_ENDPOINT_UPLOAD + filename + '.' + fileExtenstion + ':/content',
        headers: {
            'Authorization': 'Bearer ' + token, //the token is a variable which holds the token
            'content-type': fileContentType
        },
        data: byteArray
    }

    return await axios(config)
        .then(response => {
            return response.data;
        })
        .catch(error => {
            return error;
        });
}

async function postListItem(token: string, body: any): Promise<any> {
    let volljaehrig: boolean = body.age == "yes" ? true : false;
    let noImpfAusweis: boolean = body.noimpfausweis == "yes" ? true : false;

    let essgewohnheiten: string[] = body['essgewohnheiten'] ? body['essgewohnheiten'].split(";") : [];

    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT_LISTITEM,
        headers: {
            'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: {
            "fields": {
                "Title": body.vorname + ' ' + body.nachname,
                "Schar": body.schar,
                "Adresse": body.adresse + '\n' + body.plz + ' ' + body.ort,
                "Volljaehrig": volljaehrig,
                "Vormund": body.vormund,
                "Email": body.email,
                "Notfallkontakt": body.notfallkontakt,
                "Notfallnummer": body.notfallnummer,
                "Arzt": body.arzt,
                "Krankenkasse": body.kk,
                "AHV": body.ahv,
                "Krankheiten": body.krankheiten,
                "Essgewohnheiten@odata.type": "Collection(Edm.String)",
                "Essgewohnheiten": essgewohnheiten,
                "Nachricht": body.sonstiges,
                "NoImpfAusweis": noImpfAusweis
            }
        }
    }

    return await axios(config)
        .then(response => {
            return response.data;
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
                    "content": "Hallo " + body.vorname + "<br /><br /><strong>Wir freuen uns sehr, dass du am diesjährigen Pfila als Teilnehmer dabei bist!</strong><br />Die detaillierten Informationen werden zu einem späteren Zeitpunkt zugestellt.<br />Bei allfälligen Fragen wende dich bitte an: ok@pfila23.ch<br /><br />Jublastische Grüsse<br />Das Pfila23 Team"
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
            return response.data;
        })
        .catch(error => {
            return error;
        });
}