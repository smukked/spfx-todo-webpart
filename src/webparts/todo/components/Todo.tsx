import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IPersonaSharedProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import { HoverCard, HoverCardType, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';

import TodoService from '../services/TodoService';
import styles from './Todo.module.scss';
import { ITodoProps } from './ITodoProps';
import { ITask } from '../interfaces/ITask';

interface ITodoState {
  tasks: ITask[];
  loading: boolean;
  loadingTaskId: number;
}

export default class Todo extends React.Component<ITodoProps, ITodoState> {

  constructor(props) {
    super(props);

    this.state = {
      tasks: [],
      loading: true,
      loadingTaskId: 0
    };
  }

  public async componentDidMount() {
    const tasks = await TodoService.getTodos('/sites/DannyModern/Lists/Todo');
    this.setState({ tasks: tasks, loading: false });
  }

  public render(): React.ReactElement<ITodoProps> {

    const toggleTask = async (task: ITask, e: React.MouseEvent) => {
      this.setState({ loadingTaskId: task.Id });

      const updatedTask = await TodoService.updateTodo('/sites/DannyModern/Lists/Todo', task);
      if (updatedTask) {
        const newTasks = this.state.tasks.map((t) => {
          if (t.Id === updatedTask.Id) {
            t.PercentComplete = updatedTask.PercentComplete;
          }
          return t;
        });

        this.setState({ tasks: newTasks, loadingTaskId: 0 });
      }
    };

    const onRenderHoverCard = (persona: IPersonaSharedProps): JSX.Element => {
      return <div className={styles.hoverCard}>
        <div className={styles.assignedTo}>Assigned to:</div>
        <Persona
          {...persona}
          onRenderSecondaryText={(prop) => {
            return <a href={`mailto:${prop.secondaryText}`}>{prop.secondaryText}</a>;
          }}
        />
      </div>;
    };

    return (
      <div className={styles.todo}>
        <span className={styles.title}>{escape(this.props.description)}</span>
        {this.state.loading && <Spinner label="Loading tasks..." />}
        <ul className={styles.todoList}>
          {!this.state.loading && this.state.tasks.map((task, i) => {
            return <li key={i} className={styles.listItem}>
              <HoverCard
                cardDismissDelay={300}
                type={HoverCardType.plain}
                plainCardProps={{
                  onRenderPlainCard: onRenderHoverCard,
                  renderData: task.Persona
                }}>
                <h3 className={task.PercentComplete === 1 ? styles.complete : ''}>{task.Title}</h3>
              </HoverCard>
              <div className={styles.status}>
                {this.state.loadingTaskId === task.Id && <Spinner style={{ float: 'left' }} label="Updating..." labelPosition="right" />}
                {this.state.loadingTaskId !== task.Id && <Checkbox label={task.PercentComplete === 1 ? 'Open' : 'Complete'} checked={task.PercentComplete === 1} onChange={toggleTask.bind(this, task)} />}
              </div>
            </li>;
          })}
        </ul>
      </div>
    );
  }
}
