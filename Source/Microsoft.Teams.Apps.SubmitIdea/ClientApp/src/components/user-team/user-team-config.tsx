// <copyright file="user-team-config.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Text, Input, SearchIcon, Label, CloseIcon, Dropdown } from "@fluentui/react-northstar";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import Constants from "../../constants/resources";
import { ICategoryDetails } from "../models/category";
import { getAllCategories } from "../../api/category-api";
import { submitConfigCategories } from "../../api/teams-config-tab-api";

export interface IConfigState {
    url: string;
    tabName: string;
    category: string;
    loading: boolean,
    selectedCategoryList: Array<ICategoryDetails>,
    categories: Array<ICategoryDetails>,
    theme: string;
}

interface ITeamConfigDetails {
    categories: string;
    teamId: string;
}

class UserTeamConfig extends React.Component<WithTranslation, IConfigState> {
    localize: TFunction;
    userObjectId: string = "";
    teamId: string = "";
    appInsights: any;
    telemetry: string | undefined = "";

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            url: this.getBaseUrl() + "/team-ideas?theme={theme}&locale={locale}&teamId={teamId}&tenant={tid}",
            tabName: "",
            category: "",
            categories: [],
            selectedCategoryList: [],
            loading: true,
            theme: ""
        }
    }

    private getBaseUrl() {
        return window.location.origin;
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId!;
            this.teamId = context.teamId!;
            this.setState({ theme: context.theme! });
            // Initialize application insights for logging events and errors.
            this.getCategory();
        });

        microsoftTeams.settings.registerOnSaveHandler(async (saveEvent) => {
            // TODO: Call api to save category reference.
            let categoryList = this.state.selectedCategoryList.map(x => x.categoryId).join(";");
            let configureDetails: ITeamConfigDetails = {
                teamId: this.teamId,
                categories: categoryList
            }
            console.log(configureDetails)
            let response = await submitConfigCategories(configureDetails);
            if (response.status === 200 && response.data) {
                microsoftTeams.settings.setSettings({
                    entityId: "TeamIdeas",
                    contentUrl: this.state.url,
                    websiteUrl: this.state.url,
                    suggestedDisplayName: this.state.tabName,
                });
                saveEvent.notifySuccess();
            }
        });
    }


    /**
    *Get categories from API
    */
    async getCategory() {
        let category = await getAllCategories();

        if (category.status === 200 && category.data) {
            this.setState({
                categories: category.data,
            });
        }

        this.setState({
            loading: false
        });
    }

    /**
   *Sets state of tagsList by removing category using its index.
   *@param index Index of category to be deleted.
   */
    onCategoryRemoveClick = (index: number) => {
        let categories = this.state.selectedCategoryList;
        categories.splice(index, 1);
        this.setState({ selectedCategoryList: categories });
    }

    onTabNameChange = (value: string) => {

        this.setState({ tabName: value });

        if (this.state.selectedCategoryList.length > 0 && value) {
            microsoftTeams.settings.setValidityState(true);
        }
        else {
            microsoftTeams.settings.setValidityState(false);
        }
    }

    getA11ySelectionMessage = {
        onAdd: item => {
            if (item) {
                let selectedCategories = this.state.selectedCategoryList;
                let category = this.state.categories.find(category => category.categoryId === item.key);
                if (category) {
                    selectedCategories.push(category);
                    this.setState({ selectedCategoryList: selectedCategories });
                    if (selectedCategories.length > 0 && this.state.tabName) {
                        microsoftTeams.settings.setValidityState(true);
                    }
                    else {
                        microsoftTeams.settings.setValidityState(false);
                    }
                }
            }
            return "";
        },
        onRemove: item => {
            let categoryList = this.state.selectedCategoryList;
            let filterCategories = categoryList.filter(category => category.categoryId !== item.key);
            this.setState({ selectedCategoryList: filterCategories });
            if (filterCategories.length > 0 && this.state.tabName) {
                microsoftTeams.settings.setValidityState(true);
            }
            else {
                microsoftTeams.settings.setValidityState(false);
            }
            return "";
        }
    }


    public render(): JSX.Element {
        if (!this.state.loading) {
            return (
                <div className="tab-container">
                    <Flex gap="gap.smaller" column>
                        <Flex.Item>
                            <>
                                <Text content={this.localize("tabName")} />
                                <Input fluid placeholder={this.localize("tabNamePlaceholder")} value={this.state.tabName} onChange={(event: any) => this.onTabNameChange(event.target.value)} />
                            </>
                        </Flex.Item>
                        <Flex.Item>
                            <>
                                <Text styles={{ marginTop: "0.5rem" }} content={this.localize("category")} />
                                <Dropdown
                                    items={this.state.categories.map(category => {
                                        return { key: category.categoryId, header: category.categoryName }
                                    })}
                                    multiple
                                    search
                                    fluid
                                    placeholder={this.localize("categoryDropdownPlaceholder")}
                                    getA11ySelectionMessage={this.getA11ySelectionMessage}
                                />
                            </>
                        </Flex.Item>
                    </Flex>
                </div>
            );
        }
        else {
            return <></>;
        }
       
    }
}

export default withTranslation()(UserTeamConfig)