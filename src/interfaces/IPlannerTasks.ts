import { IPrimaryTextCardParameters } from "@microsoft/sp-adaptive-card-extension-base";

export interface IPlannerTasks extends IPrimaryTextCardParameters {
    id: number;
    taskName?: string;
    taskUrl?:string;
    planId?: string;
    planName?: string;
    dueDateTime?: string;
    hasDescription?: boolean;
    percentComplete?: number;
    priority?: number;
    status?:any;
    assignments?: {}
    isOverDue?: boolean;

};

