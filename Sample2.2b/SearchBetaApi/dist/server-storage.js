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
    This file wraps some of the functions of the node-persist library.
*/
const storage = require("node-persist");
class ServerStorage {
    static persist(key, data) {
        return __awaiter(this, void 0, void 0, function* () {
            yield storage.init();
            yield storage.setItem(key, data);
            console.log(key);
            console.log(data);
        });
    }
    static retrieve(key) {
        return __awaiter(this, void 0, void 0, function* () {
            yield storage.init();
            if (yield storage.getItem(key)) {
                return yield storage.getItem(key);
            }
            else {
                return null;
            }
        });
    }
    static clear() {
        return __awaiter(this, void 0, void 0, function* () {
            yield storage.init();
            yield storage.clear();
        });
    }
}
exports.ServerStorage = ServerStorage;
//# sourceMappingURL=server-storage.js.map