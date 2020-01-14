import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import TodoWebPart from '../TodoWebPart';
import { ITask, ITaskResponse } from './../interfaces/ITask';

class _TodoService {

    public async getTodos(listName: string): Promise<ITask[]> {
        const query = `${TodoWebPart.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=Id,Title,PercentComplete,AssignedTo/EMail&$expand=AssignedTo`;
        const response = await TodoWebPart.context.spHttpClient.get(query, SPHttpClient.configurations.v1);
        
        if (response.ok) {
            const data: ITaskResponse = await response.json();
            return data.value;
        } else {
            console.error(response.status, response.statusText);
            return [];
        }

    }

    public async updateTodo(listName: string, task: ITask): Promise<ITask> {

        const query = `${TodoWebPart.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${task.Id})`;

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

}
const TodoService = new _TodoService();
export default TodoService;