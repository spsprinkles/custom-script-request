import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            TitleFieldDisplayName: "Site Url",
            CustomFields: [
                {
                    name: "Owners",
                    title: "Owners",
                    type: Helper.SPCfgFieldType.User,
                    multi: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleOnly,
                    required: true
                } as Helper.IFieldInfoUser,
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "New",
                    indexed: true,
                    required: true,
                    showInNewForm: false,
                    showInEditForm: false,
                    choices: [
                        "New", "Error", "Processed", "Completed"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "Created", "Status", "Owners"
                    ]
                }
            ]
        }
    ]
});