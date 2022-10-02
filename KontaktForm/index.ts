import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const RECAPTCHA = process.env["recaptchaCodev3"]
const APP_ID = process.env["APP_ID"];
const APP_SECRET = process.env["APP_SECRET"];
const TENANT_ID = process.env["TENANT_ID"];
const TEAMS_WEBHOOK = process.env["TEAMS_WEBHOOK"];

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
    } else {
        delete req.body["g-recaptcha-response"];
    }

    try {
        // Set Default Header for Axios Requests
        axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';
        let token = await getToken();
        let response = await sendMail(token, req.body);
    
        await sendToTeams(req.body);
    
        context.log(response)
    
        context.res = {
          // status: 200, /* Defaults to 200 */
          body: req.body
        };
      } catch(e) {
        context.log("failed");
        context.res = {
          status: 500
        }
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


async function sendToTeams(body: any) {
    let config: AxiosRequestConfig = {
      method: 'post',
      url: TEAMS_WEBHOOK,
      headers: {
        'ContentType': 'Application/Json'
      },
      data: {
        "type": "message",
        "attachments": [
          {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
              "type": "AdaptiveCard",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": body.name,
                          "wrap": true,
                          "size": "Medium",
                          "weight": "Bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": body.email,
                          "wrap": true,
                          "weight": "Lighter"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "TextBlock",
                  "text": body.message,
                  "wrap": true
                }
              ],
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "version": "1.3"
            }
          }
        ]
      }
    }
  
    await axios(config);
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
          "subject": "Kontaktaufnahme",
          "body": {
            "contentType": "Text",
            "content": "Hallo " + body.name + "\n\nVielen Dank für deine Kontaktaufnahme! Wir werden uns so schnell wie möglich wieder bei dir melden.\n\nJublastische Grüsse\nDas Pfila23 Team"
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
        console.log(response.data);
        return response.data.value;
      })
      .catch(error => {
        console.log(error);
      });
  }