import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'CustomScriptRequestWebPartStrings';

export interface ICustomScriptRequestWebPartProps {
  azure_function_url: string;
}

// Reference the solution
import "main-lib";
declare const CustomScriptRequest: {
  appDescription: string;
  render: (el: HTMLElement, context: WebPartContext, azureFunctionUrl: string) => void;
  updateTheme: (theme: IReadonlyTheme) => void;
};

export default class CustomScriptRequestWebPart extends BaseClientSideWebPart<ICustomScriptRequestWebPartProps> {

  public render(): void {
    // Render the solution
    CustomScriptRequest.render(this.domElement, this.context, this.properties.azure_function_url);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    CustomScriptRequest.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Settings:",
              groupFields: [
                PropertyPaneTextField('azure_function_url', {
                  label: strings.AzureFunctionUrlFieldLabel,
                  description: strings.AzureFunctionUrlFieldDescription
                })
              ]
            }
          ]
        },
        {
          groups: [
            {
              groupName: "About this app:",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: CustomScriptRequest.appDescription
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('supportLink', {
                  href: "https://www.spsprinkles.com/",
                  text: "SharePoint Sprinkles",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/custom-script-request",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
