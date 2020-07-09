// <copyright file="upvote-idea-dialog.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Button, Flex, Text, TrashCanIcon, ItemLayout, Image, Provider, Label, Loader } from "@fluentui/react-northstar";
import { CloseIcon } from "@fluentui/react-icons-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { IDiscoverPost } from "../card-view/idea-wrapper-page";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { IdeaEntity } from "../models/idea";
import UserAvatar from "../curator-team/user-avatar";
import { generateColor } from "../../helpers/helper";
import "../../styles/edit-dialog.css";
import "../../styles/card.css";
import { getIdea } from "../../api/idea-api";
let moment = require('moment');

interface IIdeaDialogContentProps extends WithTranslation {
    cardDetails: IDiscoverPost
    onVoteClick: () => void;
    changeDialogOpenState: (isOpen: boolean) => void;
}

interface IIdeaDialogContentState {
    idea: IdeaEntity | undefined,
    loading: boolean,
    submitLoading: boolean,
    isEditDialogOpen: boolean,
    theme: string;
}

class UpvoteIdeaDialogContent extends React.Component<IIdeaDialogContentProps, IIdeaDialogContentState> {
    localize: TFunction;
    teamId = "";
    constructor(props: any) {
        super(props);

        this.localize = this.props.t;
        this.state = {
            loading: true,
            idea: undefined,
            submitLoading: false,
            isEditDialogOpen: false,
            theme: ""
        }
    }

    componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            this.teamId = context.teamId!;
            this.setState({ theme: context.theme!});
            this.getIdea(this.props.cardDetails?.ideaId);
        });
    }

    /**
   *Get idea details from API
   */
    async getIdea(ideaId: string) {
        let idea = await getIdea(ideaId!);
        if (idea.status === 200 && idea.data) {
            this.setState({ idea: idea.data });
        }
        this.setState({
            loading: false
        });
    }

	/**
	*Close the dialog and pass back card properties to parent component.
	*/
    onSubmitClick = async () =>
    {
        this.props.onVoteClick();
        this.props.changeDialogOpenState(false);
    }


	/**
	* Renders the component
	*/
    public render(): JSX.Element {
        if (!this.state.loading) {
            return (
                <Provider className="dialog-provider-wrapper">
                    <Flex>
                        <Flex.Item grow>
                            <ItemLayout
                                className="app-name-container"
                                media={<Image className="app-logo-container" src="/Artifacts/applicationLogo.png" />}
                                header={<Text content={this.localize("dialogTitleAppName")} weight="bold" />}
                                content={<Text content={this.localize("viewIdeaTitle")} weight="semibold" size="small" />}
                            />
                        </Flex.Item>
                        <CloseIcon className="icon-hover close-icon" onClick={() => this.props.changeDialogOpenState(false)} />
                    </Flex>
                    <Flex>
                        <div className="dialog-body">
                            {this.state.idea && <Flex column gap="gap.small">
                                <Text size="largest" weight="bold" content={this.state.idea.title} />
                                <div className="upvote-count">
                                    <Text size="largest" weight="bold" content={this.state.idea.totalVotes} /><br />
                                    <Text content={this.localize("UpvotesText")} />
                                </div>
                                <Flex className="margin-subcontent" vAlign="center"><UserAvatar avatarColor={generateColor()}
                                    showFullName={true} postType={this.state.idea.category!}
                                    content={this.state.idea.createdByName!} title={this.state.idea.title!} />
                                    &nbsp;<Text className="author-name" content={this.localize("ideaPostedOnText", { time: moment(new Date(this.state.idea.createdDate!)).format("llll") })} />
                                </Flex>
                                <Text content={this.localize("SynopsisText")} weight="bold" />
                                <Text content={this.state.idea.description} />
                                <Text content={this.localize("supportingDocumentsTitle")} weight="bold" />
                                <div className="documents-area">
                                    {this.state.idea.documentLinks && JSON.parse(this.state.idea.documentLinks).map((document) => <Text className="title-text" content={document} onClick={() => window.open(document, "_blank")} />)}
                                </div>
                                <Flex>
                                    <Flex.Item size="size.half">
                                        <Flex column gap="gap.small">
                                            <Text content="Tags: " weight="bold" />
                                            <div>
                                                {this.state.idea.tags ?.split(";") ?.map((tag, index) => <Label circular styles={{ marginRight: "0.2rem" }}
                                                    key={index} content={tag} />)}
                                            </div>
                                        </Flex>
                                    </Flex.Item>
                                    <Flex.Item size="size.half">
                                        <Flex column gap="gap.small">
                                            <Text weight="bold" content={this.localize("category")} />
                                            <Text content={this.state.idea.category} />
                                        </Flex>
                                    </Flex.Item>
                                </Flex>

                            </Flex>}
                        </div>
                    </Flex>
                    <Flex className="dialog-footer-wrapper">
                        <Flex gap="gap.smaller" className="dialog-footer input-fields-margin-between-add-post">
                            <div></div>
                            <Flex.Item push>
                                <Button content={this.props.cardDetails.isVotedByUser === true ? this.localize("unlikeButtonText") : this.localize("UpvoteButtonText")}
                                    primary loading={this.state.submitLoading} disabled={this.state.submitLoading} onClick={this.onSubmitClick} />
                            </Flex.Item>

                        </Flex>
                    </Flex>
                </Provider>
            );
        }
        else {
            return <Loader />
        }
    }
}

export default withTranslation()(UpvoteIdeaDialogContent)