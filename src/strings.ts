import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "custom-script-request",
    GlobalVariable: "CustomScriptRequest",
    Lists: {
        Main: "Custom Script Requests"
    },
    ProjectName: "Custom Script Requester",
    ProjectDescription: "A tool to trigger a flow to enable the custom script feature on a site collection.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;