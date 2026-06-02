import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import "@pnp/sp/views";
import { IFieldList} from "../models/models";
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, FieldUserSelectionMode, UrlFieldFormatType } from "@pnp/sp/fields";
import { IView } from "@pnp/sp/views";



const LOG_SOURCE = '🔴 Config Service';

export class ConfigService {
  private LOG_SOURCE: string = '🔴 ConfigService';
    private _sp!: SPFI;

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
    } catch (err) {
      console.error(`${LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async _configList(
    listName: string,
    listDescription: string,
    listFields:IFieldList[]): Promise<boolean> {
    let retVal = false;
    try {
      const list = await this._sp.web.lists.ensure(listName, listDescription, 100, false, { OnQuickLaunch: false });
      if (list.created) {
        await this._sp.web.lists.getByTitle(listName).fields.getByTitle("Title").update({ Hidden: true });
        await this._configSiteColumns(listFields, listName);
        retVal = true;
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE}:(_configList) - ${err}`);
    }
    return retVal;
  }

  private async _configSiteColumns(fieldList: IFieldList[], listName: string): Promise<boolean> {
    let retVal = false;
    try {
      const webInfo = await this._sp.web();
      const groupName = "_" + webInfo.Title;
      const list = await this._sp.web.lists.getByTitle(listName);
      for (let i = 0; i < fieldList.length; i++) {
        if (fieldList[i].props.FieldTypeKind === 2) {

          await list.fields.addText(fieldList[i].internalName, {
            Group: groupName,
            Title: fieldList[i].displayName
          });


        } else if (fieldList[i].props.FieldTypeKind === 3) {

          await list.fields.createFieldAsXml(
            `<Field Type="Note" Name="${fieldList[i].internalName}" DisplayName="${fieldList[i].internalName}" Required="FALSE" RichText="${fieldList[i].props.richText}" RichTextMode="FullHtml" Group="${groupName}" />`
          );
          await list.fields.getByInternalNameOrTitle(fieldList[i].internalName).update({ Title: fieldList[i].displayName });

        } else if (fieldList[i].props.FieldTypeKind === 4) {

          await list.fields.addDateTime(fieldList[i].internalName, {
            DisplayFormat: DateTimeFieldFormatType.DateOnly,
            DateTimeCalendarType: CalendarType.Gregorian,
            FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled,
            Group: groupName,
            Title: fieldList[i].displayName
          });

        } else if (fieldList[i].props.FieldTypeKind === 6) {

          await list.fields.addChoice(fieldList[i].internalName, {
            Choices: fieldList[i].props.choices ?? [],
            EditFormat: fieldList[i].props.editFormat,
            Group: groupName,
            Title: fieldList[i].displayName,
          });

        } else if (fieldList[i].props.FieldTypeKind === 8) {

          await list.fields.addBoolean(fieldList[i].internalName, { Group: groupName, Title: fieldList[i].displayName });

        } else if (fieldList[i].props.FieldTypeKind === 9) {

          await list.fields.addNumber(fieldList[i].internalName, {
            MinimumValue: fieldList[i].props.minValue,
            MaximumValue: fieldList[i].props.maxValue,
            Group: groupName,
            Title: fieldList[i].displayName
          });

        } else if (fieldList[i].props.FieldTypeKind === 10) {

          await list.fields.addCurrency(fieldList[i].internalName, {
            MinimumValue: fieldList[i].props.minValue,
            MaximumValue: fieldList[i].props.maxValue,
            CurrencyLocaleId: fieldList[i].props.localID,
            Group: groupName,
            Title: fieldList[i].displayName
          });

        } else if (fieldList[i].props.FieldTypeKind === 11) {

          await list.fields.addUrl(fieldList[i].internalName, {
            DisplayFormat: UrlFieldFormatType.Hyperlink,
            Group: groupName,
            Title: fieldList[i].displayName
          });

        } else if (fieldList[i].props.FieldTypeKind === 12) {

          await list.fields.addNumber(fieldList[i].internalName, { Group: groupName, Title: fieldList[i].displayName });

        } else if (fieldList[i].props.FieldTypeKind === 15) {

          await list.fields.addMultiChoice(fieldList[i].internalName, {
            Choices: fieldList[i].props.choices ?? [],
            FillInChoice: false,
            Group: groupName,
            Title: fieldList[i].displayName
          });

        } else if (fieldList[i].props.FieldTypeKind === 20) {

          await list.fields.addUser(fieldList[i].internalName, { SelectionMode: FieldUserSelectionMode.PeopleOnly, Group: groupName, Title: fieldList[i].displayName });

        }else if (fieldList[i].props.FieldTypeKind === 99) {

          await list.fields.addImageField(fieldList[i].internalName,{ Group: groupName, Title: fieldList[i].displayName });

        }
      }

      const view: IView = await list.defaultView;
      await view.fields.removeAll();
      for (let i = 0; i < fieldList.length; i++) {
        await view.fields.add(fieldList[i].internalName);
      }
      retVal = true;
    } catch (err) {
      console.error(`${this.LOG_SOURCE}:(_configSiteColumns) - ${err}`);
    }
    return retVal;
  }



}