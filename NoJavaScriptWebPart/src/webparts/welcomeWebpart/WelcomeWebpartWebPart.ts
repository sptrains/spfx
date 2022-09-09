import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  ThemeProvider, 
  ThemeChangedEventArgs, 
  IReadonlyTheme } from '@microsoft/sp-component-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WelcomeWebpartWebPart.module.scss';
import * as strings from 'WelcomeWebpartWebPartStrings';


export interface IWelcomeWebpartWebPartProps {
  title: string;
  messagestyle: string;
  textalignment: string;
  showtimebasedmessage: boolean;
  morningmessage: string;
  afternoonmessage: string;
  afternoonbegintime: number;
  eveningmessage: string;
  eveningbegintime: number;
  message: string;
  showname: string;
  showfirstname: boolean;
}

export default class WelcomeWebpartWebPart extends BaseClientSideWebPart<IWelcomeWebpartWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      this._themeProvider = this.context.serviceScope.consume(
        ThemeProvider.serviceKey
      );

      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();

      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(
        this,
        this._handleThemeChangedEvent
      );
    });
  }

  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // this.domElement.innerHTML = `
    // <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
    //   <div class="${styles.welcome}">
    //     <!--<img alt="Test" src="${require('./assets/welcome-dark.png')}"/>
    //     <h2>Well, ${escape(this.context.pageContext.user.displayName)}!</h2>-->
    //     <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
    //     <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
    //     <div>${this._environmentMessage}</div>
    //   </div>
    //   <div>
    //     <h3>Welcome to SharePoint Framework!</h3>        
    //     <p>
    //     The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
    //     </p>
    //     <h4>Learn more about SPFx development:</h4>
    //       <ul class="${styles.links}">
    //         <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
    //         <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
    //       </ul>
    //   </div>
    // </section>`;
    const { semanticColors }: IReadonlyTheme = this._themeVariant;

    let message = this.properties.message;

    if (this.properties.showtimebasedmessage) {
      const today: Date = new Date();
      if (today.getHours() >= this.properties.eveningbegintime) {
        message = this.properties.eveningmessage;
      }
      if (
        today.getHours() >= this.properties.afternoonbegintime &&
        today.getHours() <= this.properties.eveningbegintime
      ) {
        message = this.properties.afternoonmessage;
      }
      if (today.getHours() < this.properties.afternoonbegintime) {
        message = this.properties.morningmessage;
      }
    }
    const nameparts = this.context.pageContext.user.displayName.split(" ");
    const textalign =
      this.properties.textalignment === "left"
        ? styles.left
        : this.properties.textalignment === "right"
        ? styles.right
        : styles.center;
        
        let name = "";

        switch (this.properties.showname) {
          case "full": {
            name = this.context.pageContext.user.displayName;
            break;
          }
          case "first": {
            name = nameparts[0];
            break;
          }
        }

    const messagecontent = `<${this.properties.messagestyle}		
		style='color: ${semanticColors.bodyText}'
	  >
    ${message} ${name}		
	  </${this.properties.messagestyle}>`;
    this.domElement.innerHTML = `
	  <div class=${styles.left}>${this.properties.title}</div>
	  ${messagecontent}
	  </div>`;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
    let messagefields = [];
    if(this.properties.showtimebasedmessage){
      messagefields.push(
        PropertyPaneTextField("morningmessage",{
          label: strings.MorningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneTextField("afternoonmessage",{
          label: strings.AfternoonMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneTextField("eveningmessage",{
          label: strings.EveningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneSlider("afternoonbegintime", {
          label: strings.AfternoonBeginTimeLabel,
          min: 12,
          max: 15,
        })
      );
      messagefields.push(
        PropertyPaneSlider("eveningbegintime", {
          label: strings.EveningBeginTimeLabel,
          min: 16,
          max: 19,
        })
      );
    }else{
      messagefields.push(
        PropertyPaneTextField("message", {
          label: strings.MessageLabel,
        })
      );
    }
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.NamePropertiesGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleLabel,
                }),
                PropertyPaneChoiceGroup("showname", {
                  label: strings.ShowNameLabel,
                  options: [
                    {
                      key: "full",
                      text: "Full name",
                    },
                    {
                      key: "first",
                      text: "First name only",
                    },
                    {
                      key: "none",
                      text: "No name",
                    },
                  ],
                }),
                PropertyPaneToggle("showtimebasedmessage", {
                  label: strings.ShowTimeBasedMessageLabel,
                }),
                ...messagefields,
              ],
            },
            {
              groupName: strings.StylePropertiesGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("messagestyle", {
                  label: strings.MessageStyleLabel,
                  options: [
                    {
                      key: "h1",
                      text: "H1",
                    },
                    {
                      key: "h2",
                      text: "H2",
                    },
                    {
                      key: "h3",
                      text: "H3",
                    },
                    {
                      key: "h4",
                      text: "H4",
                    },
                    {
                      key: "p",
                      text: "P",
                    },
                  ],
                }),
                PropertyPaneChoiceGroup("textalignment", {
                  label: strings.TextAlignmentLabel,
                  options: [
                    {
                      key: "left",
                      text: "Left",
                      iconProps: {
                        officeFabricIconFontName: "AlignLeft",
                      },
                    },
                    {
                      key: "centre",
                      text: "Center",
                      iconProps: {
                        officeFabricIconFontName: "AlignCenter",
                      },
                    },
                    {
                      key: "right",
                      text: "Right",
                      iconProps: {
                        officeFabricIconFontName: "AlignRight",
                      },
                    },
                  ],
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}