import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { mountApp, type IAppHandle } from './app/app';

export interface IAssetTrackerWebPartProps {
  title: string;
  orgName: string;
  listPrefix: string;
  theme: string;
}

export default class AssetTrackerWebPart extends BaseClientSideWebPart<IAssetTrackerWebPartProps> {
  private _app: IAppHandle | undefined;

  public render(): void {
    if (this._app) {
      this._app.destroy();
    }
    const p = this.properties;
    this._app = mountApp(this.domElement, {
      webUrl: this.context.pageContext.web.absoluteUrl,
      spHttpClient: this.context.spHttpClient,
      spConfig: SPHttpClient.configurations.v1,
      listPrefix: p.listPrefix === undefined ? 'IT ' : p.listPrefix,
      theme: p.theme || 'light',
      title: (p.title || '').trim() || 'IT Asset Tracker',
      orgName: (p.orgName || '').trim(),
      version: this.manifest.version,
      userName: this.context.pageContext.user.displayName
    });
  }

  protected onDispose(): void {
    if (this._app) {
      this._app.destroy();
      this._app = undefined;
    }
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Settings for the IT Asset Tracker. Changes apply when you select Apply.' },
          groups: [
            {
              groupName: 'Header',
              groupFields: [
                PropertyPaneTextField('title', { label: 'Title', description: 'Shown at the top left. Default: IT Asset Tracker.' }),
                PropertyPaneTextField('orgName', { label: 'Organization name', description: 'Shown under the title and used in the Copilot planning prompt. Optional.' })
              ]
            },
            {
              groupName: 'Data',
              groupFields: [
                PropertyPaneTextField('listPrefix', {
                  label: 'List name prefix',
                  description: 'The lists are found by name, for example "IT Assets" and "IT People". Match the -Prefix used with Install-AssetTracker.ps1 (default "IT ").'
                })
              ]
            },
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneDropdown('theme', {
                  label: 'Color theme',
                  options: [
                    { key: 'light', text: 'Light' },
                    { key: 'dark', text: 'Dark' },
                    { key: 'system', text: "Match the viewer's device" }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
