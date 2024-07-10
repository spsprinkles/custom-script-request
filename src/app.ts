import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import { DataSource, IListItem } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
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
                            this.showNewForm();
                        }
                    }
                ]
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
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        jQuery(api.context[0].nTable).removeClass('no-footer');
                        jQuery(api.context[0].nTable).addClass('tbl-footer');
                        jQuery(api.context[0].nTable).addClass('table-striped');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[0, "desc"]]
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
                            // See if this request hasn't been processed
                            if (item.Status == "New") {
                                // See if this is the creator of the item or an admin
                                if (item.AuthorId == ContextInfo.userId) {
                                    // Render the action buttons
                                    Components.ButtonGroup({
                                        el,
                                        isSmall: true,
                                        buttons: [
                                            {
                                                text: "Delete",
                                                type: Components.ButtonTypes.OutlineDanger,
                                                onClick: () => {
                                                    // Show delete dialog
                                                    this.showDeleteDialog(item);
                                                }
                                            }
                                        ]
                                    });
                                }
                            }
                        }
                    }
                ]
            }
        });
    }

    // Shows the delete dialog
    private showDeleteDialog(item: IListItem) {
        // Set the header
        Modal.setHeader("Delete Request");

        // Set the body
        Modal.setBody("Are you sure you want to delete this request?");

        // Render a delete button
        Components.Button({
            el: Modal.FooterElement,
            type: Components.ButtonTypes.OutlineDanger,
            text: "Delete",
            onClick: () => {
                // Show a loading dialog
                LoadingDialog.setHeader("Deleting Request");
                LoadingDialog.setBody("This will close after the request is removed...");
                LoadingDialog.show();

                // Delete the item
                item.delete().execute(() => {
                    // Hide the dialog
                    LoadingDialog.hide();
                }, () => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Show an error dialog
                });
            }
        });

        // Show the dialog
        Modal.show();
    }

    // Shows the new form
    private showNewForm() {
        // Show the new form
        DataSource.List.newForm({
            onValidation: (values, isValid) => {
                // See if the form values have been set
                if (isValid) {
                    // Return a promise
                    return new Promise(resolve => {
                        let ctrlSiteUrl = DataSource.List.EditForm.getControl("Title");

                        // Show a loading dialog
                        LoadingDialog.setHeader("Getting Site Information");
                        LoadingDialog.setBody("Validating the site url...");
                        LoadingDialog.show();

                        // Get the site information
                        Web(values["Title"]).execute(
                            web => {
                                // Web exists, update the site url to be absolute
                                values["Title"] = web.Url;

                                // Hide the dialog
                                LoadingDialog.hide();

                                // Resolve the request
                                resolve(true);
                            },

                            (ex) => {
                                // See if it's a permission issue
                                if (ex.status == 403) {
                                    // User doesn't have access to site
                                    ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error the site exists, but you do not have permissions to it."
                                    });
                                }
                                // Else, the site doesn't exist
                                else {
                                    // Site doesn't exist
                                    ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error the site doesn't exist. Please check the url and try again."
                                    });
                                }

                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Resolve the request
                                resolve(false);
                            }
                        );
                    });
                }

                // Return the flag by default
                return isValid;
            },
            onUpdate: (item: IListItem) => {
                // Refresh the data
                DataSource.refresh(item.Id).then(() => {
                    // Refresh the table
                    this._dashboard.refresh(DataSource.ListItems);
                });
            }
        });
    }
}