import * as React from "react";
// import { Provider, Flex, Text, Button, Header } from "@fluentui/react";
import {
  Provider,
  Flex,
  Text,
  Button,
  Header,
  ThemePrepared,
  themes,
  Input
} from "@fluentui/react";
import TeamsBaseComponent, {
  ITeamsBaseComponentProps,
  ITeamsBaseComponentState
} from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the configMathTabTab React component
 */
export interface IConfigMathTabState extends ITeamsBaseComponentState {
  teamsTheme: ThemePrepared;
  mathOperator?: string;
  operandA: number;
  operandB: number;
  result: string;
}

/**
 * Properties for the configMathTabTab React component
 */
export interface IConfigMathTabProps extends ITeamsBaseComponentProps {}

/**
 * Implementation of the ConfigMathTab content page
 */
export class ConfigMathTab extends TeamsBaseComponent<
  IConfigMathTabProps,
  IConfigMathTabState
> {
  public componentWillMount() {
    // this.updateTheme(this.getQueryVariable("theme"));
    // this.setState({
    //   fontSize: this.pageFontSize()
    // });
    this.updateComponentTheme(this.getQueryVariable("theme"));

    if (this.inTeams()) {
      microsoftTeams.initialize();
      // microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
      microsoftTeams.registerOnThemeChangeHandler(this.updateComponentTheme);
      microsoftTeams.getContext(context => {
        // this.setState({
        //   entityId: context.entityId
        // });
        this.setState(
          Object.assign({}, this.state, {
            mathOperator: context.entityId.replace("MathPage", "")
          })
        );
        this.updateTheme(context.theme);
      });
    } else {
      // this.setState({
      //   entityId: "This is not hosted in Microsoft Teams"
      // });
      this.setState(
        Object.assign({}, this.state, {
          mathOperator: "add"
        })
      );
    }
  }

  /**
   * The render() method to create the UI of the tab
   */
  // public render() {
  //   return (
  //     <Provider theme={this.state.theme}>
  //       <Flex
  //         fill={true}
  //         column
  //         styles={{
  //           padding: ".8rem 0 .8rem .5rem"
  //         }}
  //       >
  //         <Flex.Item>
  //           <Header content="This is your tab" />
  //         </Flex.Item>
  //         <Flex.Item>
  //           <div>
  //             <div>
  //               <Text content={this.state.entityId} />
  //             </div>
  //             <div>
  //               <Button onClick={() => alert("It worked!")}>
  //                 A sample button
  //               </Button>
  //             </div>
  //           </div>
  //         </Flex.Item>
  //         <Flex.Item
  //           styles={{
  //             padding: ".8rem 0 .8rem .5rem"
  //           }}
  //         >
  //           <Text size="smaller" content="(C) Copyright Costa Engineers" />
  //         </Flex.Item>
  //       </Flex>
  //     </Provider>
  //   );
  // }

  public render() {
    return (
      <Provider theme={this.state.teamsTheme}>
        <Flex column gap="gap.smaller">
          <Header>This is your tab</Header>
          <Text content="Enter the values to calculate" size="medium"></Text>

          <Flex gap="gap.smaller">
            <Flex.Item>
              <Flex gap="gap.smaller">
                <Flex.Item>
                  <Input
                    autoFocus
                    value={this.state.operandA}
                    onChange={this.handleOnChangedOperandA}
                  ></Input>
                </Flex.Item>
                <Text content={this.state.mathOperator}></Text>
                <Flex.Item>
                  <Input
                    value={this.state.operandB}
                    onChange={this.handleOnChangedOperandB}
                  ></Input>
                </Flex.Item>
              </Flex>
            </Flex.Item>
            <Button
              content="Calculate"
              primary
              onClick={this.handleOperandChange}
            ></Button>
            <Text content={this.state.result}></Text>
          </Flex>
          <Text content="(C) Copyright Contoso" size="smallest"></Text>
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

  private handleOnChangedOperandA = (event): void => {
    this.setState(
      Object.assign({}, this.state, { operandA: event.target.value })
    );
  }

  private handleOnChangedOperandB = (event): void => {
    this.setState(
      Object.assign({}, this.state, { operandB: event.target.value })
    );
  }

  private handleOperandChange = (): void => {
    let stringResult: string = "n/a";

    if (
      !isNaN(Number(this.state.operandA)) &&
      !isNaN(Number(this.state.operandB))
    ) {
      switch (this.state.mathOperator) {
        case "add":
          stringResult = (
            Number(this.state.operandA) + Number(this.state.operandB)
          ).toString();
          break;
        case "subtract":
          stringResult = (
            Number(this.state.operandA) - Number(this.state.operandB)
          ).toString();
          break;
        case "multiply":
          stringResult = (
            Number(this.state.operandA) * Number(this.state.operandB)
          ).toString();
          break;
        case "divide":
          stringResult = (
            Number(this.state.operandA) / Number(this.state.operandB)
          ).toString();
          break;
        default:
          stringResult = "n/a";
          break;
      }
    }

    this.setState(
      Object.assign({}, this.state, {
        result: stringResult
      })
    );
  }
}
