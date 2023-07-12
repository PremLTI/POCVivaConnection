import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { PlannerTasksPropertyPane } from './PlannerTasksPropertyPane';
import { update, get } from '@microsoft/sp-lodash-subset';
import { IPlannerTasks } from '../../interfaces/IPlannerTasks';
import { GraphServiceInstance } from '../../services/graphservice';
import Utilities from '../../common/utilities';
import { PlannerTasksTypesView } from './plannerTaskTypesView/PlannerTasksTypesView';
import { PlannerTasksListView } from './plannerTaskListView/PlannerTasksListView';
import * as strings from 'PlannerTasksAdaptiveCardExtensionStrings';

export interface IPlannerTasksAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  plan: string;
}

export interface IPlannerTasksAdaptiveCardExtensionState {
  tasks: IPlannerTasks[];
  selectedTasksType: string;
  totalDueTasks?: number;
  totalUrgentTasks?:number;
  totalOverDueTasks?: number;
  plan?: string;
}


const PLANNERTASKS_MAIN_CARDVIEW: string = 'PLANNERTASKS_MAIN_CARDVIEW';
export const PLANNERTASKS_TYPES_QUICKVIEW: string = 'PLANNERTASKS_TYPES_QUICKVIEW';
export const PLANNERTASKS_LIST_QUICKVIEW: string = 'PLANNERTASKS_List_QUICKVIEW';

export default class PlannerTasksAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IPlannerTasksAdaptiveCardExtensionProps,
  IPlannerTasksAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: PlannerTasksPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
      tasks: [],
      totalDueTasks: -1,
      totalOverDueTasks: -1,
      totalUrgentTasks: -1,
      selectedTasksType: "due",
      plan: this.properties.plan ? this.properties.plan : ""
    };
    GraphServiceInstance.context = this.context;
    await this.GetPlannerTasks(this.properties.plan ? this.properties.plan : "");

    this.cardNavigator.register(PLANNERTASKS_MAIN_CARDVIEW, () => new CardView());
    this.quickViewNavigator.register(PLANNERTASKS_TYPES_QUICKVIEW, () => new PlannerTasksTypesView());
    this.quickViewNavigator.register(PLANNERTASKS_LIST_QUICKVIEW, () => new PlannerTasksListView());

    return Promise.resolve();
  }

  public get title(): string {
    return this.properties.title;
  }

  public async GetPlannerTasks(planId?: string): Promise<any> {
    let tenantGuid = this.context.pageContext.aadInfo.tenantId._guid;
    let response: any[] = Utilities.IsNullOrEmpty(planId) ? await GraphServiceInstance.GetAllTasks() : await GraphServiceInstance.GetPlannerTasks(planId);
    let tasks: IPlannerTasks[] = await Promise.all(response.map(async (task: any) => {
      const planInfo = await GraphServiceInstance.GetPlanInfo(task.planId)
      return {
        id: task.id,
        taskName: task.title,
        taskUrl: `https://teams.microsoft.com/l/entity/com.microsoft.teamspace.tab.planner/mytasks?tenantId=${tenantGuid}&webUrl=https://tasks.teams.microsoft.com/teamsui/personalApp/alltasklists&context={"subEntityId":"/taskListType/smartList/SL_Planned/list/plan/${task.planId}/task/${task.id}"}`,
        planId: task.planId,
        planName: planInfo.title,
        dueDateTime: Utilities.getLocaleDateString(task.dueDateTime),
        hasDescription: task.hasDescription,
        percentComplete: task.percentComplete,
        priority: task.priority,
        status: Utilities.GetStatus(task.percentComplete),
        assignments: task.assignments,
        isOverDue: Utilities.isDateBeforeToday(task.dueDateTime),
        description: strings.Description,
        primaryText: strings.PrimaryText
      };
    }));

    this.setState({
      tasks,
      totalDueTasks: tasks.filter(t => t.isOverDue == false).length,
      totalOverDueTasks: tasks.filter(t => t.isOverDue == true).length,
      totalUrgentTasks: tasks.filter(t => t.priority == 1).length 
    });
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'PlannerTasks-property-pane'*/
      './PlannerTasksPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.PlannerTasksPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return PLANNERTASKS_MAIN_CARDVIEW;
  }


  private onChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });


    if (newValue != oldValue) {
      this.GetPlannerTasks(newValue);
    }
    // refresh web part
    this.renderCard();
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration(this.properties, this.renderCard.bind(this), this.context, this.onChange.bind(this));
  }
}
