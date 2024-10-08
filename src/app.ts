import { Dashboard } from "dattatable";
import { Components, ContextInfo, CustomIcons, CustomIconTypes } from "gd-sprest-bs";
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
        return item.AuthorId == ContextInfo.userId;
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
                onRendering: props => {
                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex align-items-center";
                    brand.appendChild(CustomIcons(CustomIconTypes.sharePoint, 44, 44));
                    brand.append("Custom Script App");
                    props.brand = brand;
                },
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
                            // Set the filter to only look at the time
                            let filterValue = moment(item["Created"]).format("llll");
                            el.setAttribute("data-filter", filterValue);

                            // See if this is the admin
                            if (Security.IsAdmin) {
                                // Display the author
                                el.innerHTML = `
                                    <span>${item.Author.Title}</span>
                                    <br/>
                                `;
                            }

                            // Use moment to set the date/time
                            el.innerHTML += `
                                <span>${filterValue}</span>
                            `;
                        }
                    },
                    {
                        name: "Title",
                        title: "Site Url"
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

                            // See if this request hasn't been processed and this is the owner
                            if (item.Status == "New" && this.isOwner(item)) {
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