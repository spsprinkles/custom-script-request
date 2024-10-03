import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import { DataSource, IListItem } from "./ds";
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

    // Method to process a request
    private processRequest(itemId: number) {
        // Show a loading dialog
        LoadingDialog.setHeader("Processing Request");
        LoadingDialog.setBody("The request is being processed...");
        LoadingDialog.show();

        // Process the request
        DataSource.processRequest(itemId).then(
            // Success
            (msg) => {
                // Refresh the item
                DataSource.refresh(itemId).then(() => {
                    // Refresh the table
                    this._dashboard.refresh(DataSource.ListItems);
                });

                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Request Processed");

                // Set the body
                Modal.setBody(msg);

                // Show the modal
                Modal.show();
            },

            // Error
            msg => {
                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Error Processing Request");

                // Set the body
                Modal.setBody(msg);

                // Refresh the item
                DataSource.refresh(itemId).then(() => {
                    // Refresh the table
                    this._dashboard.refresh(DataSource.ListItems);

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Show the modal
                    Modal.show();
                });
            }
        );
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
                                            this.showDeleteDialog(item);
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
                                        this.processRequest(item.Id);
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
                    // Refresh the data
                    DataSource.refresh().then(() => {
                        // Refresh the dashboard
                        this._dashboard.refresh(DataSource.ListItems);

                        // Hide the dialogs
                        LoadingDialog.hide();
                        Modal.hide();
                    });
                }, () => {
                    // Hide the dialogs
                    LoadingDialog.hide();
                    Modal.hide();
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
            onCreateEditForm: props => {
                props.onControlRendering = (ctrl, fld) => {
                    // See if this is the owners field
                    if (fld.InternalName == "Owners") {
                        // Default the owners to the current user
                        ctrl.value = [ContextInfo.userId];
                    }
                }

                // Return the properties
                return props;
            },
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

                                web.SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
                                    // Ensure this is an SCA
                                    if (user.IsSiteAdmin) {
                                        // Hide the dialog
                                        LoadingDialog.hide();

                                        // Resolve the request
                                        resolve(true);
                                    } else {
                                        // Site doesn't exist
                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                            isValid: false,
                                            invalidMessage: "Site exists, but you are not the administrator. Please have the site administrator submit the request."
                                        });

                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Resolve the request
                                        resolve(false);
                                    }
                                })
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
                // See if the azure function is enabled
                if (DataSource.AzureFunctionEnabled) {
                    // Process the request
                    this.processRequest(item.Id);
                } else {
                    // Refresh the data
                    DataSource.refresh(item.Id).then(() => {
                        // Refresh the table
                        this._dashboard.refresh(DataSource.ListItems);
                    });
                }
            }
        });
    }
}