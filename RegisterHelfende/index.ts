import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getToken } from "../Common/Token";
import axios, { AxiosRequestConfig } from 'axios';
import { validateRECAP } from "../Common/Recaptcha";

const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;
const LIST_ID = process.env.HELFENDE_LIST_ID;

const MS_GRAPH_ENDPOINT_LISTITEM = 'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/lists/' + LIST_ID + '/items';
const MS_GRAPH_ENDPOINT_UPLOAD = 'https://graph.microsoft.com/v1.0/drives/' + DRIVE_ID + '/root:/';
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
        if (req.body.signature && req.body.signature.startsWith("data:")) {
            let signatureFileUpload = await uploadFile(token, req.body.signature, req.body.vorname.replace(/[^\x00-\x7F]/g, "") + '-' + req.body.nachname.replace(/[^\x00-\x7F]/g, "") + '-' + unixTime + '-signature')
            context.log("Signatur Upload: ", signatureFileUpload);
            if (signatureFileUpload.status == 201 || signatureFileUpload.status == 200) {
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
            return response;
        })
        .catch(error => {
            return error;
        });
}

async function postListItem(token: string, body: any): Promise<any> {
    let volljaehrig: boolean = body.age == "yes" ? true : false;
    let slrg: boolean = body.slrg == "yes" ? true : false;
    let auto: boolean = body.auto == "yes" ? true : false;
    let noImpfAusweis: boolean = body.noimpfausweis == "yes" ? true : false;

    let datesHelfende: string[] = body['pfila-dates'] ? body['pfila-dates'].split(";") : [];

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
                "Geburtstag": body.geburtstag,
                "Schar": body.schar,
                "Kurs": body.kurs,
                "SLRG": slrg,
                "Auto": auto,
                "Adresse": body.adresse + '\n' + body.plz + ' ' + body.ort,
                "Volljaehrig": volljaehrig,
                "Vormund": body.vormund,
                "Email": body.email,
                "Notfallkontakt": body.notfallkontakt,
                "Notfallnummer": body.notfallnummer,
                "Essgewohnheiten@odata.type": "Collection(Edm.String)",
                "Essgewohnheiten": essgewohnheiten,
                "AndereEssstoerungen": body.essstoerungen,
                "TShirt": body["shirt-size"],
                "Nachricht": body.sonstiges,
                "Schichten@odata.type": "Collection(Edm.String)",
                "Schichten": datesHelfende,
                "Anzahl": body.helpcount
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
                    "content": "Liebe*r " + body.vorname + "<br /><br /><strong>Vielen Dank fürs Anmelden als Helfende*r am Pfila 23.</strong><br />Wir werden uns bis Ende April mit den definitiven Schichten bei dir melden. Falls du Fragen oder Anmerkungen hast oder sich etwas ändert, melde dich bei Lois: luisa.fornasiero@jublaost.ch.<br />Hier noch eine kleine Info zum Zeltplatz: Es gibt die Möglichkeit, auf dem Pfilagelände zu übernachten. Dazu musst du aber dein eigenes Zelt mitnehmen. Da der Platz leider etwas begrenzt ist, kannst du nur da schlafen, wenn du an zwei Tagen einen Einsatz hast. Für Verpflegung ist selbstverständlich gesorgt.<br />Bei allgemeinen Fragen wende dich bitte an: ok@pfila23.ch<br /><br />Wir freuen uns und liebe Grüsse<br /><br />Lois und Silja<br />Ressort Helfende"
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