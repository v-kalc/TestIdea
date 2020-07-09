// <copyright file="teams-config-tab-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";
import { getBaseUrl } from '../configVariables';

let baseAxiosUrl = getBaseUrl() + '/api';

/**
* Post config tags for discover tab
* @param postContent Categories to be saved
*/
export const submitConfigCategories = async (postContent: any): Promise<any> => {
    let url = `${baseAxiosUrl}/teamcategory`;
    return await axios.post(url, postContent);
}

/**
* Get preferences tags for configure preferences
* @param teamId Team Id for which configured tags needs to be fetched
*/
export const getConfigCategories = async (teamId: string): Promise<any> => {
    let url = `${baseAxiosUrl}/teamcategory?teamId=${teamId}`;
    return await axios.get(url);
}

export const getConfigTags = async (teamId: string): Promise<any> => {
    let url = `${baseAxiosUrl}/teamcategory?teamId=${teamId}`;
    return await axios.get(url);
}

export const submitConfigTags = async (postContent: any): Promise<any> => {
    let url = `${baseAxiosUrl}/teamcategory`;
    return await axios.post(url, postContent);
}