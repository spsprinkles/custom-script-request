import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    AuthorId: number;
    Owners: { results: { Title: string; Id: number; }[] }
    OwnersId: { results: number[]; }
    Status: string;
}

/**
 * Data Source
 */
export class DataSource {
    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

    // List
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }
    private static loadList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                itemQuery: {
                    Expand: ["Owners"],
                    OrderBy: ["Title"],
                    GetAllItems: true,
                    Top: 5000,
                    Select: ["*", "Owners/Id", "Owners/Title"]
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the status filters
                    this.loadStatusFilters();

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters() {
        let items: Components.ICheckboxGroupItem[] = [];

        // Parse the choices
        let fld: Types.SP.FieldChoice = this.List.getField("Status");
        for (let i = 0; i < fld.Choices.results.length; i++) {
            // Add an item
            items.push({
                label: fld.Choices.results[i],
                type: Components.CheckboxGroupTypes.Switch
            });
        }

        // Set the filters and resolve the promise
        this._statusFilters = items;
    }

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return Promise.all([
            // Init the security
            Security.init(),
            // Load the list
            this.loadList()
        ]);
    }

    // Refreshes the list data
    static refresh(itemId?: number): PromiseLike<IListItem | IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if an item id exists
            if (itemId > 0) {
                // Refresh the item
                DataSource.List.refreshItem(itemId).then(resolve, reject);
            } else {
                // Refresh the data
                DataSource.List.refresh().then(resolve, reject);
            }
        });
    }
}