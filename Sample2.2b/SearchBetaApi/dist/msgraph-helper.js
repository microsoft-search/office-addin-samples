"use strict";
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/
const odata_helper_1 = require("./odata-helper");
class MSGraphHelper {
    // If any part of queryParamsSegment comes from user input,
    // be sure that it is sanitized so that it cannot be used in
    // a Response header injection attack.
    static getGraphData(accessToken, apiURLsegment, queryParamsSegment) {
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            const oData = yield odata_helper_1.ODataHelper.getData(accessToken, this.domain, apiURLsegment, this.versionURLsegment, queryParamsSegment);
            resolve(oData);
        }));
    }
}
MSGraphHelper.domain = "graph.microsoft.com";
MSGraphHelper.versionURLsegment = "/v1.0";
exports.MSGraphHelper = MSGraphHelper;
//# sourceMappingURL=msgraph-helper.js.map