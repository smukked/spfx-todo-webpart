import { IPersonaSharedProps } from "office-ui-fabric-react/lib/Persona";

interface IAssignedTo {
    EMail: string;
    FirstName: string;
    LastName: string;
}

export interface ITask {
    Id: number;
    Title: string;
    PercentComplete: number;
    AssignedTo: IAssignedTo;
    Persona: IPersonaSharedProps;
}

export interface ITaskResponse {
    value: ITask[];
}