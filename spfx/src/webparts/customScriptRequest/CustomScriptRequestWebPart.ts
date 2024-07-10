import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ICustomScriptRequestWebPartProps { }

// Reference the solution
import "main-lib";
declare const CustomScriptRequest: {
  render: (el: HTMLElement, context: WebPartContext) => void;
  updateTheme: (theme: IReadonlyTheme) => void;
};

export default class CustomScriptRequestWebPart extends BaseClientSideWebPart<ICustomScriptRequestWebPartProps> {

  public render(): void {
    // Render the solution
    CustomScriptRequest.render(this.domElement, this.context);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    CustomScriptRequest.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
