import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './SeitenbreiteAusnutzenWebPart.module.scss';
import * as strings from 'SeitenbreiteAusnutzenWebPartStrings';

export interface ISeitenbreiteAusnutzenWebPartProps {
  description: string;
  remark: string;
}

export default class SeitenbreiteAusnutzenWebPart extends BaseClientSideWebPart<ISeitenbreiteAusnutzenWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    if (this.isPageInEditMode()) {
      this.domElement.innerHTML = `
      <section class="${styles.seitenbreiteAusnutzen} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <div>Dieses Webpart dient dazu, die volle Seitenbreite des Browserfensters auszunutzen.
          <br>
          Wenn Sie den Bearbeitungsmodus verlassen, verschwindet diese Schrift.${this._environmentMessage}</div>
        </div>
      </section>`;
    }
    else {
      this.domElement.innerHTML = '<!-- Volle Seitennbreite-->';
    }
    this.removeMaxWidth();
  }

  // This function returns true, if the Modern Page is in Edit Mode
  private isPageInEditMode(): boolean {
    return this.displayMode === DisplayMode.Edit;
  }

  private removeMaxWidth(): void {
    const divs = document.querySelectorAll('.CanvasSection');
    divs.forEach(div => {
      const parent: HTMLDivElement = div.parentElement as HTMLDivElement;
      parent.style.maxWidth = 'none';
    });
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('remark', {
                  label: "Bemerkunng"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
