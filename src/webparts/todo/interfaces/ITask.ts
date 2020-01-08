interface IAssignedTo {
    EMail: string;
}

export interface ITask {
    Id: number;
    Title: string;
    PercentComplete: number;
    AssignedTo: IAssignedTo;
}

export interface ITaskResponse {
    value: ITask[];
}