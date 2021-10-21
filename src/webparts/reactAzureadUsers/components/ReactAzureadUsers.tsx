import * as React from 'react';
import { IReactAzureadUsersProps } from './IReactAzureadUsersProps';
import { IReactAzureadUsersState } from './IReactAzureadUsersState';
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import {  PeoplePicker } from '@microsoft/mgt-react';

export default class ReactAzureadUsers extends React.Component<IReactAzureadUsersProps, IReactAzureadUsersState> {

  constructor(props: IReactAzureadUsersProps) {
    super(props);
    this.state = {
      groupId: '',
      memers: []
    }
  }

  public onInit(): Promise<void> {
    graph.setup({
      spfxContext: this.context
    })
    return Promise.resolve();
  }

  public componentDidMount() {
    this.getAzureGroupId()
  }

  public async getAzureGroupId() {

    let groupName = "Employees"
    let getGroupIDUrl = `groups?$select=id,displayName&$filter=displayName eq '${groupName}'`;

    if (!this.props.graphClient) {
      return;
    }

    this.props.graphClient
      .api(getGroupIDUrl)
      .version("v1.0")
      .get((err: any, res: any): void => {
        if (err) {
          console.log("Getting error in retrieving azure group =>", err)
        }
        if (res) {
          if (res && res.value.length) {
            console.log(res.value);
            this.setState({
              groupId: res.value[0].id
            })
          }
          this.getGroupMembers(this.state.groupId);
        }
      });
  }

  public getGroupMembers(id: string) {
    if (!this.props.graphClient) {
      return;
    }
    let getGroupMembersUrl = `groups/${id}/members`;
    this.props.graphClient
      .api(getGroupMembersUrl)
      .version("v1.0")
      .get((err: any, res: any): void => {
        if (err) {
          console.log("Getting error in retrieving group members =>", err)
        }
        if (res) {
          if (res && res.value.length) {
            console.log(res.value);
            this.setState({
              memers: res.value
            })
          }
        }
      });
  }

  public render(): React.ReactElement<IReactAzureadUsersProps> {
    return (
      <React.Fragment>
        <h2>Azure AD Users</h2>
        <PeoplePicker
          people={this.state.memers}
          showMax={3}
        />
      </React.Fragment >
    );
  }
}
