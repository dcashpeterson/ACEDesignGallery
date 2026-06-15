import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/site-users';
import '@pnp/sp/site-groups';
import { IFacilitiesMaintenanceRequest, IRequestFormDraft, FacilitiesMaintenanceFields } from '../models/models';
import { Lists } from '../models/enums';
import { ConfigService } from './ConfigService';

const LIST_DESCRIPTION = 'Facilities and maintenance request tracking with multi-stage workflow';

export class FacilitiesMaintenanceService {
  private LOG_SOURCE: string = '🔧 FacilitiesMaintenanceService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      const created = await this._configService._configList(
        Lists.FACILITIESMAINTENANCE,
        LIST_DESCRIPTION,
        FacilitiesMaintenanceFields
      );
      if (created) {
        await this._seedSampleData();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getUserRole(maintenanceGroupName: string, managerGroupName: string): Promise<'manager' | 'maintenance' | 'user'> {
    try {
      const groups = await this._sp.web.currentUser.groups();
      const names = groups.map((g: { Title: string }) => g.Title);
      if (names.indexOf(managerGroupName) !== -1) return 'manager';
      if (names.indexOf(maintenanceGroupName) !== -1) return 'maintenance';
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getUserRole) - ${err}`);
    }
    return 'user';
  }

  public async getMyRequests(authorId: number): Promise<IFacilitiesMaintenanceRequest[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.FACILITIESMAINTENANCE)
        .items.select(
          'Id', 'Title', 'Description', 'Category', 'FacilityLocation', 'Priority', 'RequestStatus',
          'RequestedBy/Title', 'RequestedById',
          'RequestedDate', 'MaintenanceNotes', 'MaintenanceUpdatedDate',
          'AssignedMaintenanceTo/Title', 'AssignedMaintenanceToId',
          'AssignedManager/Title', 'AssignedManagerId',
          'ManagerNotes', 'CompletedDate'
        )
        .expand('RequestedBy', 'AssignedMaintenanceTo', 'AssignedManager')
        .filter(`RequestedById eq ${authorId}`)
        .orderBy('RequestedDate', false)
        .top(500)();

      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getMyRequests) - ${err}`);
      return [];
    }
  }

  public async getMaintenanceQueue(assignedToId: number): Promise<IFacilitiesMaintenanceRequest[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.FACILITIESMAINTENANCE)
        .items.select(
          'Id', 'Title', 'Description', 'Category', 'FacilityLocation', 'Priority', 'RequestStatus',
          'RequestedBy/Title', 'RequestedById',
          'RequestedDate', 'MaintenanceNotes', 'MaintenanceUpdatedDate',
          'AssignedMaintenanceTo/Title', 'AssignedMaintenanceToId',
          'AssignedManager/Title', 'AssignedManagerId',
          'ManagerNotes', 'CompletedDate'
        )
        .expand('RequestedBy', 'AssignedMaintenanceTo', 'AssignedManager')
        .filter(`AssignedMaintenanceToId eq ${assignedToId} and RequestStatus ne 'Completed' and RequestStatus ne 'Cancelled'`)
        .orderBy('Priority', false)
        .top(500)();

      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getMaintenanceQueue) - ${err}`);
      return [];
    }
  }

  public async getManagerQueue(managerId: number): Promise<IFacilitiesMaintenanceRequest[]> {
    try {
      const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();

      const [assignedItems, overdueItems] = await Promise.all([
        this._sp.web.lists
          .getByTitle(Lists.FACILITIESMAINTENANCE)
          .items.select(
            'Id', 'Title', 'Description', 'Category', 'FacilityLocation', 'Priority', 'RequestStatus',
            'RequestedBy/Title', 'RequestedById',
            'RequestedDate', 'MaintenanceNotes', 'MaintenanceUpdatedDate',
            'AssignedMaintenanceTo/Title', 'AssignedMaintenanceToId',
            'AssignedManager/Title', 'AssignedManagerId',
            'ManagerNotes', 'CompletedDate'
          )
          .expand('RequestedBy', 'AssignedMaintenanceTo', 'AssignedManager')
          .filter(`AssignedManagerId eq ${managerId} and RequestStatus ne 'Completed' and RequestStatus ne 'Cancelled'`)
          .top(500)(),

        this._sp.web.lists
          .getByTitle(Lists.FACILITIESMAINTENANCE)
          .items.select(
            'Id', 'Title', 'Description', 'Category', 'FacilityLocation', 'Priority', 'RequestStatus',
            'RequestedBy/Title', 'RequestedById',
            'RequestedDate', 'MaintenanceNotes', 'MaintenanceUpdatedDate',
            'AssignedMaintenanceTo/Title', 'AssignedMaintenanceToId',
            'AssignedManager/Title', 'AssignedManagerId',
            'ManagerNotes', 'CompletedDate'
          )
          .expand('RequestedBy', 'AssignedMaintenanceTo', 'AssignedManager')
          .filter(`RequestStatus eq 'New' and RequestedDate le datetime'${sevenDaysAgo}'`)
          .top(500)()
      ]);

      const seen = new Set<number>();
      const merged: any[] = [];
      for (const item of [...assignedItems, ...overdueItems]) {
        if (!seen.has(item.Id)) {
          seen.add(item.Id);
          merged.push(item);
        }
      }

      return merged.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getManagerQueue) - ${err}`);
      return [];
    }
  }

  public async createRequest(draft: IRequestFormDraft): Promise<string> {
    try {
      const today = new Date().toISOString().split('T')[0];
      const result = await this._sp.web.lists.getByTitle(Lists.FACILITIESMAINTENANCE).items.add({
        Title: draft.title,
        Description: draft.description,
        Category: draft.category,
        FacilityLocation: draft.location,
        Priority: draft.priority,
        RequestStatus: 'New',
        RequestedById: draft.requestedById,
        RequestedDate: today
      });
      return String(result.Id);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (createRequest) - ${err}`);
      throw err;
    }
  }

  public async updateStatus(id: string, status: string, maintenanceNotes: string): Promise<void> {
    try {
      const today = new Date().toISOString().split('T')[0];
      await this._sp.web.lists.getByTitle(Lists.FACILITIESMAINTENANCE).items.getById(Number(id)).update({
        RequestStatus: status,
        MaintenanceNotes: maintenanceNotes,
        MaintenanceUpdatedDate: today
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (updateStatus) - ${err}`);
      throw err;
    }
  }

  public async validateCompletion(id: string, managerNotes: string, approved: boolean): Promise<void> {
    try {
      const today = new Date().toISOString().split('T')[0];
      await this._sp.web.lists.getByTitle(Lists.FACILITIESMAINTENANCE).items.getById(Number(id)).update({
        RequestStatus: approved ? 'Completed' : 'In Progress',
        ManagerNotes: managerNotes,
        ...(approved ? { CompletedDate: today } : {})
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (validateCompletion) - ${err}`);
      throw err;
    }
  }

  private _mapItem(item: any): IFacilitiesMaintenanceRequest {
    return {
      id: String(item.Id),
      title: item.Title ?? '',
      description: item.Description ?? '',
      category: item.Category ?? '',
      location: item.FacilityLocation ?? '',
      priority: item.Priority ?? 'Medium',
      status: item.RequestStatus ?? 'New',
      requestedByName: item.RequestedBy?.Title ?? '',
      requestedById: item.RequestedById ?? 0,
      requestedDate: item.RequestedDate ?? '',
      assignedToName: item.AssignedMaintenanceTo?.Title ?? '',
      assignedToId: item.AssignedMaintenanceToId ?? 0,
      maintenanceNotes: item.MaintenanceNotes ?? '',
      maintenanceUpdatedDate: item.MaintenanceUpdatedDate ?? '',
      assignedManagerName: item.AssignedManager?.Title ?? '',
      assignedManagerId: item.AssignedManagerId ?? 0,
      managerNotes: item.ManagerNotes ?? '',
      completedDate: item.CompletedDate ?? ''
    };
  }

  private async _seedSampleData(): Promise<void> {
    try {
      const list = this._sp.web.lists.getByTitle(Lists.FACILITIESMAINTENANCE);
      const today = new Date().toISOString().split('T')[0];
      const tenDaysAgo = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

      const samples = [
        { Title: 'HVAC not cooling in Conference Room B', Description: 'The air conditioning unit in Conference Room B stopped working. Temperature is rising.', Category: 'HVAC', FacilityLocation: 'Building 1 - Conference Room B', Priority: 'High', RequestStatus: 'New', RequestedDate: today },
        { Title: 'Broken light fixture in hallway', Description: 'Fluorescent light fixture flickering and making buzzing sound in the main hallway near the elevator.', Category: 'Electrical', FacilityLocation: 'Building 1 - Main Hallway', Priority: 'Medium', RequestStatus: 'In Progress', RequestedDate: tenDaysAgo, MaintenanceNotes: 'Ordered replacement fixture, arriving Thursday.' },
        { Title: 'Leaking faucet in restroom', Description: 'The faucet in the 2nd floor men\'s restroom has been dripping continuously.', Category: 'Plumbing', FacilityLocation: 'Building 1 - 2F Restroom', Priority: 'Low', RequestStatus: 'Completed', RequestedDate: tenDaysAgo, CompletedDate: today },
        { Title: 'Wet floor near water fountain', Description: 'Water fountain on the 3rd floor appears to be leaking onto the floor creating a slip hazard.', Category: 'Safety Hazard', FacilityLocation: 'Building 2 - 3F Water Fountain', Priority: 'Critical', RequestStatus: 'New', RequestedDate: tenDaysAgo }
      ];

      for (const sample of samples) {
        await list.items.add(sample);
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_seedSampleData) - ${err}`);
    }
  }
}
