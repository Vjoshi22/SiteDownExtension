import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import * as $ from "jquery";
import { ISiteConfigList } from "../components/ISiteConfigProps";
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import {
    SPHttpClient, ISPHttpClientOptions,
    SPHttpClientConfiguration, SPHttpClientResponse,
    HttpClientResponse
} from "@microsoft/sp-http";


export interface IreactSiteProps {
    context: ApplicationCustomizerContext;
}
export interface IreactSiteState {
    Title: string;
    SiteStatus: string;
    Image: string;
    userMatched: boolean;
    Key:string;
    Value:string;
    GroupPermission:string;
    Category:string;
    MessageToShow:string;
}

var listGUID = '893599b8-df17-4aa6-b302-53e6d5dfeff2';
var itemID = '1';
var loggedInUser;
var ownersGroup;
var matchedResult;

export default class SiteDownPage extends React.Component<IreactSiteProps, IreactSiteState> {
    constructor(props: IreactSiteProps) {
        super(props)
        this.state = {
            SiteStatus: '',
            Title: '',
            Image: '',
            userMatched: false,
            Key:'',
            Value:'',
            GroupPermission:'',
            Category:'',
            MessageToShow:''
        }
    }
    public componentDidMount() {

        //initially hiding the Site Down Ribbion
        $("#customMessage_Top").hide();

        //fetching the configurations from the List
        this._getConfiguration();
        /**
         * Checking the current logged in user
         * Check the LoggedIn user in PT- Admin Group
         * Hide show according to the user Present in the above group or not
         */
        this._getCurrentUser();

    }
    public componentWillMount() {

    }
    public render(): JSX.Element {
        return (
            <div id="customMessage_Top">
                <h4>{this.state.MessageToShow}</h4>
            </div>
        )
    }
    private _getConfiguration() {
        const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists('${listGUID}')/items?$select=*&$filter=Key eq 'IS SITE DOWN'`;
        return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }).then((response: SPHttpClientResponse) => {
                return response.json();
            })
            .then((item): void => {
                this.setState({
                    Key: item.value[0].Key,
                    Value: item.value[0].Value,
                    GroupPermission: item.value[0].GroupPermission,
                    Category: item.value[0].Category,
                    MessageToShow: item.value[0].MessageToShow
                })
                this._getCurrentUser();
                // this._checkSiteStatus();
            })
    }
    //check the Status and applied the Overlay
    private _checkSiteStatus() {
        if (this.state.Value == "Yes" && !this.state.userMatched) {
            $("#customMessage_Top").show();
            $('#spSiteHeader').nextAll().remove();
            $('.Files-mainColumn').remove();
            $('#customMessage_Top').css({
                "position": "absolute",
                "top": "50%",
                "left": "50%",
                "font-size": "25px",
                "color": "black",
                "transform": "translate(-50%,-50%)",
                "-ms-transform": "translate(-50%,-50%)",
                "width":"80%"
            });
        } else if (this.state.Value == "Yes" && this.state.userMatched) {
            $("#customMessage_Top").show();
            $('.Files-mainColumn').show();
            $('#spSiteHeader').nextAll().show();
        } else {
            $("#customMessage_Top").remove();
        }
    }
    private _getCurrentUser() {
        //get the current user context
        this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
            SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                response.json().then((responseJSON: any) => {
                    //console.log(responseJSON);
                    console.log(responseJSON.DisplayName)
                    loggedInUser = responseJSON.DisplayName;
                }).then(grp => {
                    this._checkUserInGroup();
                })
            })
    }
    private _checkUserInGroup() {
        //Get the user from a group
        this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyname('PM - Admin ')/users?select=Title`,
            SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                console.log("logging response" + response);
                response.json().then((rjson: any) => {
                    console.log("Group Name")
                    console.log(JSON.stringify(rjson.value))
                    rjson.value.forEach((getUser: any) => {
                        console.log(getUser.Title);
                        if (ownersGroup == "") {
                            ownersGroup = `${getUser.Title}`;
                        }
                        else {
                            ownersGroup = ownersGroup + "@!" + `${getUser.Title}`;
                        }
                    });
                    matchedResult = ownersGroup.match(loggedInUser);
                    if (matchedResult != null) {
                        this.setState({
                            userMatched: true
                        })
                        this._checkSiteStatus();
                    } else {
                        this._checkSiteStatus();
                    }
                });
            });
    }
}