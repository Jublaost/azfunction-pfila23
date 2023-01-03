import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getToken } from "../Common/Token";
import axios, { AxiosRequestConfig } from 'axios';
import { validateRECAP } from "../Common/Recaptcha";

const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;
const LIST_ID = process.env.TN_LIST_ID;

const MS_GRAPH_ENDPOINT_LISTITEM = 'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/lists/' + LIST_ID + '/items';
const MS_GRAPH_ENDPOINT_UPLOAD = 'https://graph.microsoft.com/v1.0/drives/' + DRIVE_ID + '/root:/';
const MS_GRAPH_ENDPOINT_SENDMAIL = 'https://graph.microsoft.com/v1.0/users/ok@pfila23.ch/sendMail';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    let token = await getToken();
    //context.log("Token: ", token)
    context.log("Body: ", req.body)

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
        if (req.body.signature) {
            let signatureFileUpload = await uploadFile(token, req.body.signature, req.body.vorname + '-' + req.body.nachname + '-signature')
            context.log("Signatur Upload: ", signatureFileUpload);
            if (signatureFileUpload.status == 200) {
                context.log(signatureFileUpload);
            } else {
                context.log("Error: ", signatureFileUpload)
                context.res = {
                    status: 500,
                    body: "server error"
                };
                return
            }
        }
        if (req.body.impfausweis) {
            let impfausweisFileUpload = await uploadFile(token, req.body.impfausweis, req.body.vorname + '-' + req.body.nachname + '-impfausweis')
            context.log("Impfausweis Upload: ", impfausweisFileUpload);
            if (impfausweisFileUpload.status == 200) {
                context.log(impfausweisFileUpload);
            } else {
                context.log("Error: ", impfausweisFileUpload)
                context.res = {
                    status: 500,
                    body: "server error"
                };
                return
            }
        }

        let response = await postListItem(token, req.body);
        context.log("Status: ", response.status);
        if (response.status == 201) {
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
                "Geschlecht": body.gender,
                "Schar": body.schar,
                "Adresse": body.adresse + '\n' + body.plz + ' ' + body.ort,
                "Vormund": body.vormund,
                "Email": body.email,
                "Notfallkontakt": body.notfallkontakt,
                "Notfallnummer": body.notfallnummer,
                "Arzt": body.hausarzt,
                "Unfallversicherung": body.unfallversicherung,
                "Krankenkasse": body.kk,
                "AHV": body.ahv,
                "Krankheiten": body.krankheiten,
                "Allergien": body.allergien,
                "Essgewohnheiten@odata.type": "Collection(Edm.String)",
                "Essgewohnheiten": essgewohnheiten,
                "AndereEssstoerungen": body.essstoerungen,
                "TShirt": body["shirt-size"],
                "Nachricht": body.sonstiges,
                "NoImpfAusweis": noImpfAusweis
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
            return response;
        })
        .catch(error => {
            return error;
        });
}