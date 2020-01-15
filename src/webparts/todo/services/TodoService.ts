import { ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import { getInitials } from '@uifabric/utilities';
import { IPersonaSharedProps, PersonaSize } from 'office-ui-fabric-react/lib/Persona';

import TodoWebPart from '../TodoWebPart';
import { ITask, ITaskResponse } from './../interfaces/ITask';

class _TodoService {

    public async getTodos(listName: string): Promise<ITask[]> {
        const select = 'Id,Title,PercentComplete,AssignedTo/EMail,AssignedTo/FirstName,AssignedTo/LastName';
        const expand = 'AssignedTo';
        // const query = `${TodoWebPart.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=${select}&$expand=${expand}`;
        const query = `${TodoWebPart.context.pageContext.web.absoluteUrl}/_api/web/GetList('${listName}')/items?$select=${select}&$expand=${expand}`;
        const response = await TodoWebPart.context.spHttpClient.get(query, SPHttpClient.configurations.v1);

        if (response.ok) {
            const data: ITaskResponse = await response.json();
            const tasks = this.addPersona(data.value);
            return tasks;
        } else {
            console.error(response.status, response.statusText);
            return [];
        }
    }

    public async updateTodo(listName: string, task: ITask): Promise<ITask> {

        const query = `${TodoWebPart.context.pageContext.web.absoluteUrl}/_api/web/GetList('${listName}')/items(${task.Id})`;

        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Accept', 'application/json;odata=nometadata');
        requestHeaders.append('Content-type', 'application/json;odata=nometadata');
        requestHeaders.append('odata-version', '');
        requestHeaders.append('IF-MATCH', '*');
        requestHeaders.append('X-HTTP-Method', 'MERGE');

        task.PercentComplete = (task.PercentComplete && task.PercentComplete === 1) ? 0 : 1;

        const body: string = JSON.stringify({
            PercentComplete: task.PercentComplete
        });

        const httpClientOptions: ISPHttpClientOptions = {
            headers: requestHeaders,
            body: body
        };

        const response = await TodoWebPart.context.spHttpClient.post(query, SPHttpClient.configurations.v1, httpClientOptions);
        if (response.ok) {
            return task;
        } else {
            return null;
        }
    }

    private addPersona(tasks: ITask[]): ITask[] {
        tasks.map(t => {
            const fullName = `${t.AssignedTo.FirstName} ${t.AssignedTo.LastName}`;
            const persona: IPersonaSharedProps = {
                id: t.AssignedTo.EMail,
                text: fullName,
                imageUrl: `/_vti_bin/DelveApi.ashx/people/profileimage?size=S&userId=${t.AssignedTo.EMail}`,
                imageAlt: fullName,
                imageInitials: getInitials(fullName, false),
                secondaryText: t.AssignedTo.EMail,
                size: PersonaSize.size40
            };

            t.Persona = persona;
        });
        return tasks;
    }

}
const TodoService = new _TodoService();
export default TodoService;