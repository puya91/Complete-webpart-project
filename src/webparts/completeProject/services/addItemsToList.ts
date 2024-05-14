import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IRiskEventStates } from "../models/riskEventStates";
import { SPFI } from '@pnp/sp';
import { getSP } from "../../../pnpjsConfig";
import { IFilePickerResult } from "@pnp/spfx-controls-react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/sites";
import "@pnp/sp/attachments";
import { IItem } from "@pnp/sp/items/types";


const attachDocument = async (_sp: SPFI, itemId: number, documentPeriodListId: string, fileItem: IFilePickerResult) => {
    fileItem.downloadFileContent()
        .then(async result => {
            const item: IItem = _sp.web.lists.getById(documentPeriodListId).items.getById(itemId);
            await item.attachmentFiles.add(fileItem.fileName, result);
        })
}

export const addItemsToList = async (context: WebPartContext, states: IRiskEventStates) => {
    const sp: SPFI = getSP();
    const requestListGuid = '48db0c6b-7b64-499f-8e4a-035499aef8f2';
    const documentsListGuid = '59eed830-55d8-4736-92ad-4244ef1a2eec';

    try{
        const addedElement = await sp.web.lists.getById(requestListGuid).items.add({
            RiskTitle: states.riskTitle,
            CreatedBy: states.createdBy,
            CreationDate: states.creationDate,
            Business: states.business,
            Country: states.country,
            RiskDate: states.riskDate,
            AssignedTo: states.assignee,
            Notes: states.notes,
            State: states.state,
            ContainsDocuments: states.containsDocuments
        })

        const elementId = addedElement.ID;

        await sp.web.lists.getById(requestListGuid).items.getById(elementId).update({
            Title: `RiskEventRequestsId-${elementId}`
        });

        if (states.riskReport != undefined) {
            const elementTitle = `RiskEventRequestsId-${elementId}`;
            const documentsListPath = `${context.pageContext.web.absoluteUrl}/Lists/RiskEventDocumentsList`;
            
            await sp.web.lists.getById(documentsListGuid).rootFolder.addSubFolderUsingPath(elementTitle);
    
            states.riskReport.forEach(async (element: IFilePickerResult) => {
                await sp.web.lists.getById(documentsListGuid).addValidateUpdateItemUsingPath([
                    {
                        FieldName: "Title",
                        FieldValue: `${element['fileName']}`,
                    }
                ], `${documentsListPath}/${elementTitle}`).then(result => {
                    attachDocument(sp, Number(result[1].FieldValue), documentsListGuid, element)
                });
            })
        }
    }
    catch (error) {
        console.error('Error adding items to the list:', error);
    }
}