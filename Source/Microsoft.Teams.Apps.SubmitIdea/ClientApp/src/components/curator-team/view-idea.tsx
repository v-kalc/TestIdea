// <copyright file="view-idea.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { WithTranslation, withTranslation } from "react-i18next";
import * as microsoftTeams from "@microsoft/teams-js";
import { Text, Flex, Provider, Label, RadioGroup, TextArea, Loader, Button, Dropdown, BanIcon, AcceptIcon } from "@fluentui/react-northstar";
import { TFunction } from "i18next";
import { IdeaEntity, ApprovalStatus } from "../models/idea";
import UserAvatar from "./user-avatar";
import { generateColor, isNullorWhiteSpace } from "../../helpers/helper";
import { ICategoryDetails } from "../models/category";
import { getAllCategories } from "../../api/category-api";
import { getIdea, updatePostContent } from "../../api/idea-api";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { createBrowserHistory } from "history";
let moment = require('moment');

interface IState {
    idea: IdeaEntity | undefined,
    loading: boolean,
    selectedStatus: number | undefined,
    selectedCategory: string | undefined,
    feedbackText: string | undefined,
    categories: Array<ICategoryDetails>,
    submitLoading: boolean,
    isCategorySelected: boolean,
    feedbackTextEmpty: boolean,
    isIdeaApprovedOrRejected: boolean;
}

const browserHistory = createBrowserHistory({ basename: "" });

class ViewIdea extends React.Component<WithTranslation, IState> {
    localize: TFunction;
    userObjectId: string | undefined = "";
    items: any;
    appInsights: any;
    telemetry: string | undefined = "";
    ideaId: string | undefined = "";

    constructor(props) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            loading: true,
            idea: undefined,
            selectedStatus: ApprovalStatus.Approved,
            selectedCategory: undefined,
            categories: [],
            feedbackText: "",
            submitLoading: false,
            isCategorySelected: false,
            feedbackTextEmpty: true,
            isIdeaApprovedOrRejected: false,
        }
        this.items = [
            {
                key: 'approve',
                label: this.localize('radioApprove'),
                value: ApprovalStatus.Approved,
            },
            {
                key: 'reject',
                label: this.localize('radioReject'),
                value: ApprovalStatus.Rejected,
            }
        ]

        let params = new URLSearchParams(window.location.search);
        this.telemetry = params.get("telemetry")!;
        this.ideaId = params.get("id")!;
    }



    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId!;

            // Initialize application insights for logging events and errors.
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            this.getCategory();
            this.getIdea();
        });
    }

    getA11SelectionMessage = {
        onAdd: item => {
            if (item) { this.setState({ selectedCategory: item, isCategorySelected: true }) };
            return "";
        },
    };

    /**
 *Get idea details from API
 */
    async getIdea() {
        this.appInsights.trackTrace({ message: `'getIdea' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let idea = await getIdea(this.ideaId!);
        if (idea.status === 200 && idea.data) {
            this.appInsights.trackTrace({ message: `'getIdea' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let category = this.state.categories.filter(row => row.categoryName === idea.data.category).shift();
            if (category === undefined) {
                this.setState({ selectedCategory: undefined });
            }
            else {
                this.setState({ selectedCategory: idea.data.category, isCategorySelected: true });
            }

            this.setState(
                {
                    loading: false,
                    idea: idea.data,
                });
        }
        else {
            this.appInsights.trackTrace({ message: `'getIdea' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            loading: false
        });
    }

    /**
  *Get categories from API
  */
    async getCategory() {
        this.appInsights.trackTrace({ message: `'getCategory' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let category = await getAllCategories();

        if (category.status === 200 && category.data) {
            this.appInsights.trackTrace({ message: `'getCategory' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            console.log(category.data);
            this.setState({
                categories: category.data,
            });
        }
        else {
            this.appInsights.trackTrace({ message: `'getCategory' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            loading: false
        });
    }

    /**
 *Approve or rejectIdea
 */
    async approveOrRejectIdea(idea: any) {
        this.appInsights.trackTrace({ message: `'approveOrRejectIdea' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let updateEntity = await updatePostContent(idea);

        if (updateEntity.status === 200 && updateEntity.data) {
            this.appInsights.trackTrace({ message: `'approveOrRejectIdea' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        else {
            this.appInsights.trackTrace({ message: `'approveOrRejectIdea' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }

        this.setState({
            loading: false,
            submitLoading: false,
            isIdeaApprovedOrRejected: true,
        });
    }

    /**
   * Handle radio group change event.
   * @param e | event
   * @param props | props
   */
    handleChange = (e: any, props: any) => {
        this.setState({ selectedStatus: props.value })
    }

    checkIfConfirmAllowed = () => {
        if (this.state.selectedCategory === undefined) {
            this.setState({ isCategorySelected: false });
            return false;
        }

        if (this.state.selectedStatus === 2 && isNullorWhiteSpace(this.state.feedbackText!)) {
            this.setState({ feedbackTextEmpty: false });
            return false;
        }

        return true;
    }

    /**
   *Returns text component containing error message for failed name field validation
   *@param {boolean} isValuePresent Indicates whether value is present
   */
    private getRequiredFieldError = (isValuePresent: boolean) => {
        if (!isValuePresent) {
            return (<Text content={this.localize('fieldRequiredMessage')} className="field-error-message" error size="medium" />);
        }

        return (<></>);
    }

    handleConfirm = () => {
        if (this.checkIfConfirmAllowed()) {
            this.setState({ submitLoading: true });
            let category = this.state.categories.filter(row => row.categoryName === this.state.selectedCategory).shift();
            let updateEntity: IdeaEntity = {
                ideaId: this.state.idea ?.ideaId,
                feedback: this.state.selectedStatus === ApprovalStatus.Rejected ? this.state.feedbackText : "",
                status: this.state.selectedStatus,
                category: this.state.selectedCategory,
                categoryId: category ?.categoryId,
                approverOrRejecterUserId: this.userObjectId,
                createdByObjectId: this.state.idea ?.createdByObjectId,
                title: this.state.idea ?.title,
                description: this.state.idea ?.description,
                documentLinks: this.state.idea ?.documentLinks,
                totalVotes: this.state.idea ?.totalVotes,
                tags: this.state.idea ?.tags,
                createdDate: this.state.idea ?.createdDate,
                createdByName: this.state.idea ?.createdByName,
                createdByUserPrincipleName: this.state.idea ?.createdByUserPrincipleName,
                updatedDate: this.state.idea ?.updatedDate,
                approvedOrRejectedByName: this.state.idea ?.approvedOrRejectedByName
            }

            this.approveOrRejectIdea(updateEntity);
        }
    }

    onFeedbackChange = (value: string) => {
        this.setState({ feedbackText: value });
    }

    /**
     * Renders the component.
     */
    public render(): JSX.Element {
        if (!this.state.loading && !this.state.isIdeaApprovedOrRejected) {
            return (
                <Provider>
                    <div className="module-container">
                        {this.state.idea && <Flex column gap="gap.small">
                            <Text size="largest" weight="bold" content={this.state.idea.title} />
                            <Flex vAlign="center"><UserAvatar avatarColor={generateColor()} showFullName={true}
                                postType={this.state.idea.category!} content={this.state.idea.createdByName!}
                                title={this.state.idea.createdByName!} />
                                &nbsp;<Text content={this.localize("ideaPostedOnText", { time: moment(new Date(this.state.idea.createdDate!)).format("llll")})} /></Flex>
                            <Text content={this.localize("synopsisTitle")} weight="bold" />
                            <TextArea className="response-text-area" value={this.state.idea.description} disabled />
                            <Flex>
                                <Flex.Item size="size.half">
                                    <Flex column gap="gap.small">
                                        <Text content={this.localize("tagsTitle")} weight="bold" />
                                        <div>
                                            {this.state.idea.tags ?.split(";").map((tag, index) => <Label circular styles={{ marginRight: "0.2rem" }} key={index} content={tag} />)}
                                        </div>
                                        <Text content={this.localize("supportingDocumentsTitle")} weight="bold" />
                                        <div className="documents-area">{this.state.idea.documentLinks && JSON.parse(this.state.idea.documentLinks).map((document) => <Text className="title-text" content={document} onClick={() => window.open(document, "_blank")} />)}</div>
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <Flex column gap="gap.small">
                                        <Text content={this.localize("category")} />
                                        <Flex.Item push>
                                            {this.getRequiredFieldError(this.state.isCategorySelected)}
                                        </Flex.Item>
                                        <Dropdown fluid
                                            items={this.state.categories.map((category) => category.categoryName)}
                                            value={this.state.selectedCategory}
                                            placeholder={this.localize("categoryPlaceholder")}
                                            getA11ySelectionMessage={this.getA11SelectionMessage}
                                        />
                                        {<RadioGroup items={this.items}
                                            defaultCheckedValue={this.state.selectedStatus}
                                            onCheckedValueChange={this.handleChange}
                                        />}
                                        {this.state.selectedStatus === ApprovalStatus.Rejected && <>
                                            {this.getRequiredFieldError(this.state.feedbackTextEmpty)}
                                            <TextArea className="reason-text-area" placeholder={this.localize("reasonForRejectionText")}
                                                value={this.state.feedbackText} onChange={(event: any) => this.onFeedbackChange(event.target.value)} /></>}

                                    </Flex>
                                </Flex.Item>
                            </Flex>

                        </Flex>}

                    </div>
                    <div className="tab-footer">
                        <Flex hAlign="end" ><Button primary disabled={this.state.submitLoading} loading={this.state.submitLoading}
                            content={this.localize("Confirm")} onClick={this.handleConfirm} /></Flex>
                    </div>
                </Provider>)
        }
        else if (this.state.isIdeaApprovedOrRejected) {
            return (
                <div className="submit-idea-success-message-container">
                    <div className="space">
                    </div>
                    <div>{this.state.selectedStatus === ApprovalStatus.Approved ? <AcceptIcon className="info-icon" size="large" /> : <BanIcon className="info-icon" size="large" />}
                    </div>
                    <div className="space"></div>
                    <Text weight="semibold"
                        content={this.state.selectedStatus === ApprovalStatus.Approved ? this.localize("approvedIdeaSuccessMessage") : this.localize("rejectedIdeaMessage")}
                        size="medium"
                    />
                    <div className="space"></div>
                </div>)
        }
        else {
            return <Loader />
        }
    }
}

export default withTranslation()(ViewIdea)