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
}

var listGUID = 'af951c71-5ca6-4bf8-b29e-55c72a25adf8';
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
            userMatched: false
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
                <h3>{this.state.Title}</h3>
            </div>
        )
    }
    private _getConfiguration() {
        const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists('${listGUID}')/items(${itemID})`;
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
                    Title: item.Title,
                    Image: item.Image.Url,
                    SiteStatus: item.SiteStatus
                })
                this._getCurrentUser();
                // this._checkSiteStatus();
            })
    }
    private _checkSiteStatus() {
        if (this.state.SiteStatus == "Down" && !this.state.userMatched) {
            $("#customMessage_Top").show();
            $('#spSiteHeader').nextAll().hide();
            $('#customMessage_Top').css({
                "position": "absolute",
                "top": "50%",
                "left": "50%",
                "font-size": "50px",
                "color": "black",
                "transform": "translate(-50%,-50%)",
                "-ms-transform": "translate(-50%,-50%)"
            });
        } else if (this.state.SiteStatus == "Down" && this.state.userMatched) {
            $("#customMessage_Top").show();
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