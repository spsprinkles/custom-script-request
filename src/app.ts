import { Dashboard } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import * as moment from "moment";
import { DataSource, IListItem } from "./ds";
import { Forms } from "./forms";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the dashboard
        this.render(el);
    }

    // Returns true/false if the user is an owner of the item
    private isOwner(item: IListItem) {
        // See if they created the request
        if (item.AuthorId == ContextInfo.userId) { return true; }

        // Parse the owners
        let ownerIDs = item.OwnersId?.results || [];
        for (let i = 0; i < ownerIDs.length; i++) {
            // See if they are an owner
            if (ownerIDs[i] == ContextInfo.userId) { return true; }
        }

        // Not the owner of the item
        return false;
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        this._dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        this._dashboard.filter(2, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create Request",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            Forms.newForm(() => {
                                // Refresh the table
                                this._dashboard.refresh(DataSource.ListItems);
                            });
                        }
                    }
                ],
                itemsEnd: Security.IsAdmin ? [
                    {
                        className: "btn-outline-light me-2",
                        text: "Settings",
                        isButton: true,
                        items: [
                            {
                                text: "Application Configuration",
                                onClick: () => {
                                    // Show the install modal
                                    InstallationModal.show(true);
                                }
                            },
                            {
                                text: "List Settings",
                                onClick: () => {
                                    // Show the list settings of the main list
                                    window.open(DataSource.List.ListSettingsUrl, "_blank");
                                }
                            }
                        ]
                    }
                ] : null
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.ListItems,
                onRendering: dtProps => {
                    // Default order
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "Created",
                        title: "Request Date",
                        onRenderCell: (el, column, item: IListItem) => {
                            // Use moment to set the date/time
                            el.innerHTML = moment(item["Created"]).format("llll");
                        }
                    },
                    {
                        name: "Title",
                        title: "Site Url"
                    },
                    {
                        name: "",
                        title: "Owners",
                        onRenderCell: (el, column, item: IListItem) => {
                            let owners = [];

                            // Parse the users
                            let users = item.Owners?.results || [];
                            for (let i = 0; i < users.length; i++) {
                                // Append the name
                                owners.push(users[i].Title);
                            }

                            // Render the owners
                            el.innerHTML = owners.join("\n<br/>\n");
                        }
                    },
                    {
                        name: "Status",
                        title: "Status"
                    },
                    {
                        name: "Actions",
                        title: "",
                        onRenderCell: (el, column, item: IListItem) => {
                            let buttons: Components.IButtonProps[] = [];

                            // See if this request hasn't been processed
                            if (item.Status == "New") {
                                // See if this is the creator of the item or an admin
                                if (this.isOwner(item)) {
                                    // Render the delete button
                                    buttons.push({
                                        text: "Delete",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Show delete dialog
                                            Forms.deleteForm(item, () => {
                                                // Refresh the table
                                                this._dashboard.refresh(DataSource.ListItems);
                                            });
                                        }
                                    });
                                }
                            }

                            // See if the request errored and the azure function is enabled
                            if (item.Status == "Error" && DataSource.AzureFunctionEnabled) {
                                // Render the retry button
                                buttons.push({
                                    text: "Retry",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Process the request
                                        Forms.processRequest(item.Id).then(() => {
                                            // Refresh the table
                                            this._dashboard.refresh(DataSource.ListItems);
                                        });
                                    }
                                });
                            }

                            // Render the buttons
                            Components.ButtonGroup({
                                el,
                                isSmall: true,
                                buttons
                            });
                        }
                    }
                ]
            }
        });
    }

}