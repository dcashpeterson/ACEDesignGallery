import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { ISafetyTip, SafetyTipsFields } from '../models/models';
import { Lists } from '../models/enums';
import { ConfigService } from './ConfigService';

const LIST_DESCRIPTION = 'Workplace safety tips by category, displayed as a daily rotating tip';

export class SafetyTipsService {
  private LOG_SOURCE = '🦺 SafetyTipsService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      const created = await this._configService._configList(Lists.SAFETYTIPS, LIST_DESCRIPTION, SafetyTipsFields);
      if (created) {
        await this._seedSampleData();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getSafetyTips(): Promise<ISafetyTip[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.SAFETYTIPS)
        .items.select('Id', 'Title', 'Summary', 'Guidance', 'Category', 'Policy_x0020_Document_x0020_URL', 'Policy_x0020_Document_x0020_Titl')
        .top(100)();

      if (items.length === 0) return this._getSampleData();

      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getSafetyTips) - ${err}`);
      return this._getSampleData();
    }
  }

  private async _seedSampleData(): Promise<void> {
    try {
      const list = this._sp.web.lists.getByTitle(Lists.SAFETYTIPS);
      const samples = this._getSampleData();
      for (const item of samples) {
        await list.items.add({
          Title: item.title,
          Summary: item.summary,
          Guidance: item.guidance,
          Category: item.category,
          Policy_x0020_Document_x0020_URL: { Url: item.policyUrl, Description: item.policyTitle },
          Policy_x0020_Document_x0020_Titl: item.policyTitle
        });
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_seedSampleData) - ${err}`);
    }
  }

  private _mapItem(item: any): ISafetyTip {
    return {
      id: String(item.Id),
      title: item.Title ?? '',
      summary: item.Summary ?? '',
      guidance: item.Guidance ?? '',
      category: item.Category ?? '',
      policyUrl: item.Policy_x0020_Document_x0020_URL?.Url ?? '',
      policyTitle: item.Policy_x0020_Document_x0020_Titl ?? 'View Policy Document'
    };
  }

  private _getSampleData(): ISafetyTip[] {
    return [
      {
        id: '1',
        title: 'Proper Lifting Technique',
        summary: 'Bend at the knees, keep the load close to your body, and never twist while lifting.',
        guidance: 'Before lifting any object, assess its weight — if it exceeds 50 lbs, ask for help or use mechanical assistance. Position yourself close to the load with your feet shoulder-width apart. Bend at your knees and hips (not your waist), keeping your back straight. Grip the object firmly and lift by straightening your legs, letting your leg muscles do the work. Keep the load close to your torso at all times, and turn by moving your feet rather than twisting your spine.',
        category: 'Ergonomics',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/ManualHandlingPolicy.pdf',
        policyTitle: 'Manual Handling & Lifting Policy'
      },
      {
        id: '2',
        title: 'Fire Extinguisher Usage',
        summary: 'Use the PASS technique — Pull, Aim, Squeeze, Sweep — on small contained fires only.',
        guidance: 'Only attempt to use a fire extinguisher if the fire is small, contained, and you have a clear exit behind you. Use the PASS method: Pull the safety pin from the handle. Aim the nozzle at the base of the fire, not the flames. Squeeze the handle to discharge the extinguisher. Sweep the nozzle from side to side, keeping it aimed at the base until the fire is out. If the fire grows or you run out of agent, evacuate immediately and call emergency services.',
        category: 'Fire Safety',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/FireSafetyPolicy.pdf',
        policyTitle: 'Fire Safety & Emergency Response Policy'
      },
      {
        id: '3',
        title: 'PPE for Chemical Handling',
        summary: 'Always wear gloves, safety goggles, and a lab coat when handling hazardous chemicals.',
        guidance: 'Before working with any chemical, consult its Safety Data Sheet (SDS) to understand the required Personal Protective Equipment. At minimum, wear chemical-resistant gloves appropriate for the substance, safety goggles with side shields, and a lab coat or chemical-resistant apron. Never eat, drink, or touch your face while handling chemicals. Work in a well-ventilated area or fume hood when required. After handling, remove PPE in the correct order to avoid contamination and wash your hands thoroughly.',
        category: 'Chemical Safety',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/HazardousSubstancesPolicy.pdf',
        policyTitle: 'Hazardous Substances & PPE Policy'
      },
      {
        id: '4',
        title: 'Emergency Exit Awareness',
        summary: 'Know the two nearest emergency exits from your workstation before an emergency occurs.',
        guidance: 'On your first day in any new workspace, identify the two closest emergency exits and mentally plan your evacuation route. Never assume the nearest exit is the main entrance — secondary exits may be closer. Do not block emergency exits or exit routes with equipment, boxes, or personal items at any time. During an evacuation, do not use elevators; always take the stairs. Gather at your designated muster point and do not re-enter the building until the all-clear is given by emergency personnel.',
        category: 'Emergency Preparedness',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/EmergencyEvacuationPolicy.pdf',
        policyTitle: 'Emergency Evacuation Policy'
      },
      {
        id: '5',
        title: 'Slip, Trip & Fall Prevention',
        summary: 'Keep walkways clear of cords, boxes, and spills to prevent the leading cause of workplace injuries.',
        guidance: 'Slips, trips, and falls are the most common cause of workplace injuries. Keep all walkways, corridors, and stairwells free of clutter, boxes, and cables. Report spills immediately and place wet floor signage until the area is dry. Secure electrical cords and cables with covers or run them along baseboards — never across walkways. Wear appropriate footwear with slip-resistant soles. When carrying large loads, ensure your path is clear and your view is unobstructed before moving.',
        category: 'General Safety',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/SlipTripFallPolicy.pdf',
        policyTitle: 'Slip, Trip & Fall Prevention Policy'
      },
      {
        id: '6',
        title: 'Computer Workstation Ergonomics',
        summary: 'Set your monitor at eye level and keep your wrists neutral to prevent repetitive strain injuries.',
        guidance: 'An ergonomically configured workstation reduces the risk of musculoskeletal disorders. Set your monitor so the top of the screen is at or slightly below eye level and at arm\'s length distance (approximately 20–24 inches). Adjust your chair so your feet rest flat on the floor and your thighs are parallel to it. Keep your keyboard and mouse at a height that allows your wrists to remain straight and neutral — not bent up or down. Take micro-breaks every 30 minutes: look at something 20 feet away for 20 seconds and stretch your hands and shoulders.',
        category: 'Ergonomics',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/WorkstationErgonomicsPolicy.pdf',
        policyTitle: 'Workstation Ergonomics Policy'
      },
      {
        id: '7',
        title: 'Hand Washing & Hygiene',
        summary: 'Wash hands for at least 20 seconds with soap to reduce the spread of illness in the workplace.',
        guidance: 'Proper hand hygiene is one of the most effective ways to prevent the spread of illness. Wash your hands with soap and water for at least 20 seconds (about the time it takes to sing "Happy Birthday" twice). Always wash after using the restroom, before and after eating, after coughing or sneezing, and after touching shared surfaces. When soap and water are unavailable, use an alcohol-based hand sanitiser with at least 60% alcohol. Avoid touching your face — particularly your eyes, nose, and mouth — with unwashed hands.',
        category: 'Health & Hygiene',
        policyUrl: 'https://contoso.sharepoint.com/sites/Safety/Policies/WorkplaceHealthPolicy.pdf',
        policyTitle: 'Workplace Health & Hygiene Policy'
      }
    ];
  }
}
