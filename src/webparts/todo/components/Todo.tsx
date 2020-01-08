import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';

import TodoService from '../services/TodoService';
import styles from './Todo.module.scss';
import { ITodoProps } from './ITodoProps';
import { ITask } from '../interfaces/ITask';

interface ITodoState {
  tasks: ITask[];
}

export default class Todo extends React.Component<ITodoProps, ITodoState> {

  constructor(props) {
    super(props);

    this.state = {
      tasks: []
    };

  }

  public async componentDidMount() {
    const tasks = await TodoService.getTodos('Todo List');
    this.setState({ tasks: tasks });
  }

  public render(): React.ReactElement<ITodoProps> {

    const toggleTask = async (task: ITask, e: React.MouseEvent) => {
      e.persist();

      const updatedTask = await TodoService.updateTodo('Todo List', task);
      if (updatedTask) {
        const newTasks = this.state.tasks.map((t) => {
          if (t.Id === updatedTask.Id) {
            t.PercentComplete = updatedTask.PercentComplete;
          }
          return t;
        });
        
        this.setState({ tasks: newTasks });
      }
    };

    return (
      <div className={styles.todo}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Todos</span>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <ul>
                {this.state.tasks.map((task, i) => {
                  return <li className={task.PercentComplete === 1 ? styles.complete : ''} key={i}>{task.Title} - {task.AssignedTo.EMail}<button onClick={toggleTask.bind(this, task)}>x</button></li>;
                })}
              </ul>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
