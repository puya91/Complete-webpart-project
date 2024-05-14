import { IFilePickerResult } from "@pnp/spfx-controls-react";

export interface IRiskEventStates {
    riskTitle: string;
    createdBy: string | null;
    creationDate: Date;
    business: string | null;
    country: string | null;
    riskDate: Date | undefined;
    assignee: string;
    containsDocuments: string;
    riskReport: IFilePickerResult[] | undefined;
    notes: string | undefined;
    state: string;
}