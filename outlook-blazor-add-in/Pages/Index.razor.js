// Copyright (c) Maarten van Stam and Eric Legault. All rights reserved. Licensed under the MIT License.

// https://github.com/justmarks/UiPathOutlook/blob/2247bb3f5f52f8ba6763374fea3f25e1ece5d03b/src/commands/commands.js

//export async function getEmailData() {
//    try {
//        console.log(`Reading mailbox item`);

//        const item = Office.context.mailbox.item;

//        if (item === null) {
//            console.error(`Index.razor.js(getEmailData): Unexpected - could not get reference to Office.context.mailbox.item`);
//            return { Subject: "No email item" };
//        }

//        console.log("Email Subject: " + item.subject);

//        var listOfAttachments = [];
//        var attachments = item.attachments;

//        console.log(`Index.razor.js(getEmailData): Counting ${item.attachments.length} attachment(s)...`);

//        if (item.attachments.length > 0) {
//            console.log("Trying to fetch attachments...", new Date());

//            try {
//                const returnedAttachments = await getAttachments(item);
//                console.log(returnedAttachments);

//                for (var i = 0; i < returnedAttachments.length; i++) {
//                    console.log(`Index.razor.js (processAttachmentsAsync2): Attachment - ` + i);
//                    console.log(`Index.razor.js (processAttachmentsAsync2): AttachmentID - ` + item.attachments[i].id);

//                    let attachmentName = item.attachments[i].name;
//                    console.log(`Index.razor.js(getEmailData): Processing attachment '${attachmentName}'...`);

//                    let attachmentType = item.attachments[i].attachmentType;
//                    let fileinline = item.attachments[i].isInline;

//                    if (attachmentType == "item") {
//                        console.log(`Index.razor.js(getEmailData): Only read file type attachments`);
//                        continue;
//                    }

//                    console.log(`Index.razor.js(getEmailData) Type: ` + attachmentType);

//                    var fileExtension = getFileExtension(attachmentName);

//                    if (fileExtension !== "") {
//                        console.log(`Index.razor.js(getEmailData): Processing file extension '${fileExtension}'...`);
//                        console.log("Attachment Value: " + returnedAttachments[i].value.content);

//                        return {
//                            AttachmentId: item.attachments[i].id,
//                            AttachmentName: attachmentName,
//                            Subject: item.subject,
//                            AttachmentType: attachmentType,
//                            Inline: fileinline,
//                            AttachmentBase64Data: returnedAttachments[i].value.content
//                        };
//                    } else {
//                        console.error(`Index.razor.js(getEmailData): Could not parse file extension for ${attachmentName}`);
//                        continue;
//                    }
//                }
//            } catch (e) {
//                console.error("Error fetching attachments:", e);
//            }
//        } else {
//            return {
//                AttachmentId: "",
//                AttachmentName: "No Attachments",
//                Subject: item.subject,
//                AttachmentType: "",
//                Inline: false,
//                AttachmentBase64Data: ""
//            };
//        }

//        // Count emails received today using EWS
//        const emailCount = await countEmailsReceivedToday();
//        console.log(`Number of emails received today: ${emailCount}`);
//        return { Subject: item.subject, EmailCount: emailCount };

//    } catch (err) {
//        console.error(`Index.razor.js(getEmailData) Catch Exception: ${err}`);
//        return { Subject: `Error: ${err.message}` };
//    }
//}
//function getFileExtension(fileName) {
//    var a = fileName.split(".");
//    if (a.length === 1 || (a[0] === "" && a.length === 2)) {
//        return "";
//    }
//    return a.pop();
//}

//async function getAttachments(item) {
//    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
//        return [];
//    }

//    const attachments = item.attachments;

//    if (!attachments) {
//        return [];
//    }

//    return promise.all(
//        attachments.map(
//            (attachment) =>
//                new promise((resolve) => {
//                    console.log(`index.razor.js (getattachments) attachment id returned: ` + attachment.id);
//                    item.getattachmentcontentasync(attachment.id, (result) => resolve(result));
//                })
//        )
//    );
//}
export async function countEmailsReceivedToday() {
    try {
        console.log("==========================  STARTED EMAIL COUNT ==========================");
        //const ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"; // URL for EWS
        //const ewsHeaders = {
        //    "Content-Type": "text/xml",
        //    "Accept": "application/xml",
        //    "Authorization": `Bearer ${Office.context.mailbox.getCallbackTokenAsync()}`
        //};

        const currentDate = new Date();
        const startOfDay = currentDate.toISOString().split('T')[0] + "T00:00:00Z";
        const endOfDay = currentDate.toISOString().split('T')[0] + "T23:59:59Z";

        const ewsRequest1 = `
             <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                <soap:Header>
                    <t:RequestServerVersion Version="Exchange2013" />
                </soap:Header>
                <soap:Body>
                    <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" Traversal="Shallow">
                        <ItemShape>
                            <t:BaseShape>IdOnly</t:BaseShape>
                        </ItemShape>
                        <IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning" />
                        <Restriction>
                            <t:And>
                                <t:IsGreaterThanOrEqualTo>
                                    <t:FieldURI FieldURI="item:DateTimeReceived" />
                                    <t:FieldURIOrConstant>
                                        <t:Constant Value="${startOfDay}" />
                                    </t:FieldURIOrConstant>
                                </t:IsGreaterThanOrEqualTo>
                                <t:IsLessThanOrEqualTo>
                                    <t:FieldURI FieldURI="item:DateTimeReceived" />
                                    <t:FieldURIOrConstant>
                                        <t:Constant Value="${endOfDay}" />
                                    </t:FieldURIOrConstant>
                                </t:IsLessThanOrEqualTo>
                            </t:And>
                        </Restriction>
                        <ParentFolderIds>
                            <t:DistinguishedFolderId Id="inbox" />
                        </ParentFolderIds>
                    </FindItem>
                </soap:Body>
            </soap:Envelope>`;

        console.log("==========================  END EMAIL COUNT ==========================");

        return new Promise((resolve, reject) => {
            Office.context.mailbox.makeEwsRequestAsync(ewsRequest1, function (asyncResult) {
                try {
                    console.log("Inside makeEwsRequestAsync callback...");
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const response = asyncResult.value;
                        //console.log("Full Response:", response);  // Log the full response here

                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");

                        const messages = xmlDoc.getElementsByTagName("t:Message");
                        const emailCount = messages.length;

                        console.log("Received Today Email count:", emailCount);
                        resolve({ Count: emailCount });  // Return email count in a resolved Promise
                    } else {
                        console.error("Request failed:", asyncResult.error.message);
                        reject(new Error(asyncResult.error.message));  // Reject if failed
                    }
                } catch (err) {
                    console.error("Error in callback:", err);
                    reject(err);  // Reject if there's a try-catch error
                }
                console.log("Exiting makeEwsRequestAsync callback...");
            });
        });

    } catch (err) {
        console.error("Error counting emails:", err);
        return 0;
    }
}


export async function countUnreadEmails() {
    try {
        console.log("==========================  STARTED UNREAD EMAIL COUNT ==========================");

        const currentDate = new Date();
        const startOfDay = currentDate.toISOString().split('T')[0] + "T00:00:00Z";
        const endOfDay = currentDate.toISOString().split('T')[0] + "T23:59:59Z";
        const ewsRequest2 = `
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
        <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" Traversal="Shallow">
            <ItemShape>
                <t:BaseShape>IdOnly</t:BaseShape>
            </ItemShape>
            <IndexedPageItemView MaxEntriesReturned="10" Offset="0" BasePoint="Beginning" />
            <Restriction>
                <t:IsEqualTo>
                    <t:FieldURI FieldURI="message:IsRead" />
                    <t:FieldURIOrConstant>
                        <t:Constant Value="false" />
                    </t:FieldURIOrConstant>
                </t:IsEqualTo>
            </Restriction>
            <ParentFolderIds>
                <t:DistinguishedFolderId Id="inbox" />
            </ParentFolderIds>
        </FindItem>
    </soap:Body>
</soap:Envelope>`;
;

        console.log("Request for unread emails:", ewsRequest2);


        console.log("==========================  END UNREAD EMAIL COUNT ==========================");

        return new Promise((resolve, reject) => {
            Office.context.mailbox.makeEwsRequestAsync(ewsRequest2, function (asyncResult) {
                try {
                    console.log("Inside makeEwsRequestAsync callback...");
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const response = asyncResult.value;

                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");

                        const messages = xmlDoc.getElementsByTagName("t:Message");
                        const unreadEmailCount = messages.length;

                        console.log("Unread Email count:", unreadEmailCount);
                        resolve({ UnCount: unreadEmailCount });  // Return unread email count in a resolved Promise
                    } else {
                        console.error("Request failed:", asyncResult.error.message);
                        reject(new Error(asyncResult.error.message));  // Reject if failed
                    }
                } catch (err) {
                    console.error("Error in callback:", err);
                    reject(err);  // Reject if there's a try-catch error
                }
                console.log("Exiting makeEwsRequestAsync callback...");
            });
        });

    } catch (err) {
        console.error("Error counting unread emails:", err);
        return 0;
    }
}


//// Index.razor.js

//export function getEmailData() {
//    return new Promise((resolve, reject) => {
//        // Assume you're using the Outlook REST API or Microsoft Graph API to get the email data
//        // Here's an example of how you might retrieve email data using Microsoft Graph:

//        Office.context.mailbox.item.subject.getAsync(function (result) {
//            if (result.status === Office.AsyncResultStatus.Succeeded) {
//                const emailData = {
//                    subject: result.value,
//                    attachmentBase64Data: "someBase64EncodedData" // Replace with actual Base64 data if needed
//                };
//                resolve(emailData);
//            } else {
//                reject('Error retrieving email data');
//            }
//        });
//    });
//}

