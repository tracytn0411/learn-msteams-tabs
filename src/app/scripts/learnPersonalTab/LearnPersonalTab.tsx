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
  Alert,
  List,
  Icon,
  Label,
  Input
} from "@fluentui/react";
import TeamsBaseComponent, {
  ITeamsBaseComponentProps,
  ITeamsBaseComponentState
} from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the learnPersonalTabTab React component
 */
export interface ILearnPersonalTabState extends ITeamsBaseComponentState {
  entityId?: string;
  teamsTheme: ThemePrepared;
  todoItems: string[];
  newTodoValue: string;
}

/**
 * Properties for the learnPersonalTabTab React component
 */
export interface ILearnPersonalTabProps extends ITeamsBaseComponentProps {}

/**
 * Implementation of the LearnPersonalTab content page
 */
export class LearnPersonalTab extends TeamsBaseComponent<
  ILearnPersonalTabProps,
  ILearnPersonalTabState
> {
  public componentWillMount() {
    // this.updateTheme(this.getQueryVariable('theme'));
    // this.setState({
    //   fontSize: this.pageFontSize()
    // });

    this.updateComponentTheme(this.getQueryVariable("theme"));
    this.setState(
      Object.assign({}, this.state, {
        todoItems: ["Submit time sheet", "Submit expense report"],
        newTodoValue: ""
      })
    );

    if (this.inTeams()) {
      microsoftTeams.initialize();
      // microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
      microsoftTeams.registerOnThemeChangeHandler(this.updateComponentTheme);
      microsoftTeams.getContext(context => {
        this.setState({
          entityId: context.entityId
        });
        this.updateTheme(context.theme);
      });
    } else {
      this.setState({
        entityId: "This is not hosted in Microsoft Teams"
      });
    }
  }

  /**
   * The render() method to create the UI of the tab
   */
  public render() {
    return (
      <Provider theme={this.state.teamsTheme}>
        <Flex column gap="gap.smaller">
          <Header>This is your tab</Header>
          <Alert
            icon="exclaimation-triangle"
            content={this.state.entityId}
            dismissible
          ></Alert>
          <Text content="There are your to-do items:" size="medium"></Text>
          <List selectable>
            {this.state.todoItems.map(todoItem => (
              <List.Item
                media={<Icon name="window-maximize outline"></Icon>}
                content={todoItem}
              ></List.Item>
            ))}
          </List>
          TODO: add new list item form here
          <Text content="(C) Copyright Tracy" size="smallest"></Text>
        </Flex>
      </Provider>
    );
  }

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

  // updates the component state to the theme that matches the currently selected Microsoft Teams client theme
  private updateComponentTheme = (teamsTheme: string = "default"): void => {
    let theme: ThemePrepared;

    switch (teamsTheme) {
      case "default":
        theme = themes.teams;
        break;
      case "dark":
        theme = themes.teamsDark;
        break;
      case "contrast":
        theme = themes.teamsHighContrast;
        break;
      default:
        theme = themes.teams;
        break;
    }
    // update the state
    this.setState(
      Object.assign({}, this.state, {
        teamsTheme: theme
      })
    );
  };
}
