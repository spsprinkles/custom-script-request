import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";


// Forms
export class Forms {
    // Shows the delete dialog
    static deleteForm(item: IListItem, onDelete: () => void) {
        // Clear the modal
        Modal.clear();

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
                        // Call the update event
                        onDelete();

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
    static newForm(onUpdate: () => void) {
        // Show the new form
        DataSource.List.newForm({
            onSetHeader: el => {
                // Set the header
                el.firstElementChild.innerHTML = "Create Request";
            },
            onFormButtonsRendering: buttons => {
                // Modify the create button's text
                buttons[0].text = "Submit";

                // Return the buttons
                return buttons;
            },
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
                        let webUrl = values["Title"];
                        Web(webUrl).query({ Expand: ["ParentWeb"] }).execute(
                            web => {
                                // See if this is a sub-web
                                if (web.ParentWeb && web.ParentWeb.Id) {
                                    // Site doesn't exist
                                    ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                        isValid: false,
                                        invalidMessage: "Detected a sub-web. Please enter the root web url."
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Resolve the request
                                    resolve(false);
                                    return;
                                }

                                // Web exists, update the site url to be absolute
                                values["Title"] = web.Url;

                                Web(web.Url).SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
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
                // Show a loading dialog
                LoadingDialog.setHeader("Refreshing Requests");
                LoadingDialog.setBody("Refreshing the data...");
                LoadingDialog.show();

                // Refresh the data
                DataSource.refresh(item.Id).then(() => {
                    // See if the azure function is enabled
                    if (DataSource.AzureFunctionEnabled) {
                        // Update the loading dialog
                        LoadingDialog.setBody("Processing the request...");

                        // Process the request
                        this.processRequest(item.Id).then(() => {
                            // Call the update event
                            onUpdate();

                            // Hide the dialog
                            LoadingDialog.hide();
                        });
                    } else {
                        // Call the update event
                        onUpdate();

                        // Hide the dialog
                        LoadingDialog.hide();
                    }
                });
            }
        });
    }

    // Method to process a request
    static processRequest(itemId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
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
                        // Hide the loading dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        resolve();
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

                    // Show the modal
                    Modal.show();

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                }
            );
        });
    }
}