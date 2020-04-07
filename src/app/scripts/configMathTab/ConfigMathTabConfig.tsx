import * as React from "react";
// import { Provider, Flex, Header, Input } from "@fluentui/react";
import {
  Provider,
  Flex,
  Header,
  Input,
  ThemePrepared,
  themes,
  DropdownProps,
  Dropdown
} from "@fluentui/react";
import TeamsBaseComponent, {
  ITeamsBaseComponentProps,
  ITeamsBaseComponentState
} from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IConfigMathTabConfigState extends ITeamsBaseComponentState {
  teamsTheme: ThemePrepared;
  mathOperator: string;
}

export interface IConfigMathTabConfigProps extends ITeamsBaseComponentProps {}

/**
 * Implementation of ConfigMathTab configuration page
 */
export class ConfigMathTabConfig extends TeamsBaseComponent<
  IConfigMathTabConfigProps,
  IConfigMathTabConfigState
> {
  public componentWillMount() {
    // this.updateTheme(this.getQueryVariable("theme"));
    // this.setState({
    // 	fontSize: this.pageFontSize()
    // });
    this.updateComponentTheme(this.getQueryVariable("theme"));

    if (this.inTeams()) {
      microsoftTeams.initialize();

      microsoftTeams.getContext((context: microsoftTeams.Context) => {
        // this.setState({
        //     value: context.entityId
        // });
        this.setState(
          Object.assign({}, this.state, {
            mathOperator: context.entityId.replace("MathPage", "")
          })
        );
        this.updateTheme(context.theme);
        this.setValidityState(true);
      });

      microsoftTeams.settings.registerOnSaveHandler(
        (saveEvent: microsoftTeams.settings.SaveEvent) => {
          // Calculate host dynamically to enable local debugging
          const host = "https://" + window.location.host;
          microsoftTeams.settings.setSettings({
            contentUrl: host + "/configMathTab/?data=",
            suggestedDisplayName: "Config Math Tab",
            removeUrl: host + "/configMathTab/remove.html",
            // entityId: this.state.value
            entityId: `${this.state.mathOperator}MathPage`
          });
          saveEvent.notifySuccess();
        }
      );
    } else {
    }
  }

  public render() {
    return (
      <Provider theme={this.state.teamsTheme}>
        <Flex gap="gap.smaller" style={{ height: "300px" }}>
          <Dropdown
            placeholder="Select the math operator"
            items={["add", "subtract", "multiply", "divide"]}
            onChange={this.handleOnSelectedChange}
          ></Dropdown>
        </Flex>
      </Provider>
    );
  }
  private updateComponentTheme = (teamsTheme: string = "default"): void => {
    let componentTheme: ThemePrepared;

    switch (teamsTheme) {
      case "default":
        componentTheme = themes.teams;
        break;
      case "dark":
        componentTheme = themes.teamsDark;
        break;
      case "contrast":
        componentTheme = themes.teamsHighContrast;
        break;
      default:
        componentTheme = themes.teams;
        break;
    }
    // update the state
    this.setState(
      Object.assign({}, this.state, {
        teamsTheme: componentTheme
      })
    );
  }

  private handleOnSelectedChange = (event, props: DropdownProps): void => {
    this.setState(
      Object.assign({}, this.state, {
        mathOperator: props.value ? props.value.toString() : "add"
      })
    );
  }

  // public render() {
  //   return (
  //     <Provider theme={this.state.theme}>
  //       <Flex fill={true}>
  //         <Flex.Item>
  //           <div>
  //             <Header content="Configure your tab" />
  //             <Input
  //               placeholder="Enter a value here"
  //               fluid
  //               clearable
  //               value={this.state.value}
  //               onChange={(e, data) => {
  //                 if (data) {
  //                   this.setState({
  //                     value: data.value
  //                   });
  //                 }
  //               }}
  //               required
  //             />
  //           </div>
  //         </Flex.Item>
  //       </Flex>
  //     </Provider>
  //   );
  // }
}
