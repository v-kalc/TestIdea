// <copyright file="categoy-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get all categories data from API
*/
export const getAllCategories = async (): Promise<any> => {

    let url = baseAxiosUrl + `/api/Category/allcategories`;
    return await axios.get(url, undefined);
}

/**
* Post category data to API
*/
export const postCategory = async (data: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/Category/category";
    return await axios.post(url, data, undefined);
}

/**
* Delete user selected category
* @param {string} categoryIds selected category ids which needs to be deleted
*/
export const deleteSelectedCategories = async (categoryIds: string): Promise<any> => {

    let url = baseAxiosUrl + `/api/Category/categories?categoryIds=${categoryIds}`;
    return await axios.delete(url);
}