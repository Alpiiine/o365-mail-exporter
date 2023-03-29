import {AuthenticationHandler, Client, ClientOptions} from "@microsoft/microsoft-graph-client";
import {MyLoggingHandler} from "./MyLoggingHandler";
import {MyHttpMessageHandler} from "./MyHttpMessageHandler";
import {MyAuthenticationProvider} from "./MyAuthenticationProvider";
import * as fs from "fs";

let authProvider = new MyAuthenticationProvider();
let authenticationHandler = new AuthenticationHandler(authProvider);
let myLoggingHandler = new MyLoggingHandler();
let myHttpMessageHandler = new MyHttpMessageHandler();

// Note: myHttpMessageHandler is the last in the chain so there is no need of setting next middleware for it.
authenticationHandler.setNext(myLoggingHandler);
myLoggingHandler.setNext(myHttpMessageHandler);

let clientOptions: ClientOptions = {
    authProvider: authProvider,
}

const client = Client.initWithMiddleware(clientOptions);

async function getMails() {
    let mailList = [];
    let mails;

    do {
        console.log("Getting mails...");
        const url = mails ? mails["@odata.nextLink"]: "/me/messages";
        console.log(`URL: ${url}`);
        mails = await client.api(url).get();
        mailList = [...mailList, ...(await parseMails(mails))];
        console.log(`Total mails: ${mailList.length}, current mails: ${mails.value.length}`);
    } while (mails["@odata.nextLink"]);

    const totalMails = mailList.length;
    console.log(`Total mails: ${totalMails}`);

    fs.writeFileSync("output/mails.json", JSON.stringify(mailList, null, 2));
}

async function parseMails(mails) {
    let mailList = [];
    for (let i = 0; i < mails.value.length; i++) {
        const mail = mails.value[i];
        if (mail.hasAttachments) {
            const attachments = await client.api(`/me/messages/${mail.id}/attachments`).get();
            for (let j = 0; j < attachments.value.length; j++) {
                const attachment = attachments.value[j];
                const attachmentSize = attachment.size;
                const attachmentType = attachment["@odata.type"];
                console.log(`Attachment Name: ${attachment.name}, size: ${attachmentSize}, type: ${attachmentType}`);

                if (attachmentType === "#microsoft.graph.referenceAttachment") {
                    // this type of attachment is stored on Onedrive or Sharepoint
                    // which the file content is not provided in the response
                    // we can get the downloadable link to the attachment by calling the following BETA API
                    const refAttachment = await client
                        .api(`/me/messages/${mail.id}/attachments/${attachment.id}`)
                        .version("beta") // this is a BETA API
                        .get();

                    // this is the downloadable link to the attachment
                    // for now, user needs to manually download the attachment
                    const refSourceUrl = refAttachment.sourceUrl;
                    mail.attachmentSourceUrl = refSourceUrl;
                    console.log(`Reference Attachment Downloadable Link: ${refSourceUrl}`);
                    continue;
                }

                // check duplicate files
                if (fs.existsSync(`output/${attachment.name}`)) {
                    // if file already exists, add last 6 characters of id to the file name
                    fs.writeFileSync("output/"
                        + attachment.id.slice(-6) + attachment.name, attachment.contentBytes, "base64");
                } else {
                    fs.writeFileSync("output/" + attachment.name, attachment.contentBytes, "base64");
                    mail.attachmentName = attachment.name;
                }
            }
        }
        mailList.push(mail);
    }

    return mailList;
}

getMails().then(r =>
    console.log("Done")
);