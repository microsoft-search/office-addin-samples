"use strict";
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
Object.defineProperty(exports, "__esModule", { value: true });
/*
    This file provides the provides error handling.
*/
class ServerError extends Error {
    /**
     * @constructor
     *
     * @param message Error message to be propagated.
    */
    constructor(name, code = 500, message, innerError) {
        super(message);
        this.code = code;
        this.message = message;
        this.innerError = innerError;
        this.name = `${name}: ${message}`;
        if (Error.captureStackTrace) {
            Error.captureStackTrace(this, this.constructor);
        }
        else {
            let error = new Error();
            if (error.stack) {
                let last_part = error.stack.match(/[^\s]+$/);
                this.stack = `\n${this.name} at ${last_part}`;
            }
        }
    }
}
exports.ServerError = ServerError;
class UnauthorizedError extends ServerError {
    constructor(message, innerError) {
        super('UnauthorizedError', 401, message, innerError);
        this.innerError = innerError;
    }
}
exports.UnauthorizedError = UnauthorizedError;
class BadRequestError extends ServerError {
    constructor(message, innerError) {
        super('BadRequestError', 400, message, innerError);
        this.innerError = innerError;
    }
}
exports.BadRequestError = BadRequestError;
class NotFoundError extends ServerError {
    constructor(message, innerError) {
        super('NotFoundError', 404, message, innerError);
        this.innerError = innerError;
    }
}
exports.NotFoundError = NotFoundError;
//# sourceMappingURL=errors.js.map