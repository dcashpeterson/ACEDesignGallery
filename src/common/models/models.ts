import { ChoiceFieldFormatType } from "@pnp/sp/fields";
import { statusType } from "./enums";

export interface IChoice {
  choice: string;
  value: string;
}
export class Choice implements IChoice {
  constructor(
    public choice: string = "",
    public value: string = ""
  ) { }
}

export interface IFieldList {
  internalName: string;
  displayName: string;
  props: { FieldTypeKind: number, choices?: string[], richText?: boolean, editFormat?: ChoiceFieldFormatType, minValue?: number, maxValue?: number, localID?: number };
}

export const LunchMenuFields: IFieldList[] = [
  { internalName: "Location", displayName: "Location", props: { FieldTypeKind: 2 } },
  { internalName: "Description", displayName: "Description", props: { FieldTypeKind: 3 } },
  { internalName: "Category", displayName: "Category", props: { FieldTypeKind: 6, choices: ["Entree", "Side", "Dessert", "Beverage", "Special"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Price", displayName: "Price", props: { FieldTypeKind: 9 } },
  { internalName: "Image", displayName: "Image", props: { FieldTypeKind: 99 } }
];

export interface IServiceHealth {
  id: string;
  title: string;
  service: string;
  statusType: statusType;
  issues?: IServiceHealthIssue[];
}

export class ServiceHealth implements IServiceHealth {
  constructor(
    public id: string,
    public title: string = "",
    public service: string = "",
    public statusType: statusType = 0 as unknown as statusType,
    public issues?: IServiceHealthIssue[]
  ) { }
}

export interface ILocation {
  name: string;
  imageUrl: string;
}

export interface ILunchMenuItem {
  id: string;
  title: string;
  location: string;
  description: string;
  category: string;
  price: number;
  image: string;
}

export interface IServiceHealthIssue {
  id: string;
  title: string;
  service: string;
  statusType: statusType;
  classification: string;
  feature: string;
  featureGroup: string;
  impactDescription: string;
  isResolved: boolean;
  origin: string;
  startDateTime: string;
  endDateTime: string;
  lastModifiedDateTime: string;
}

