// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/

import * as fs from 'fs';
import * as https from 'https';
import * as path from 'path';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as morgan from 'morgan';
import { AuthModule } from './auth';

/* Set the environment to development if not set */
const env = process.env.NODE_ENV || 'development';

/* Instantiate AuthModule to assist with JWT parsing and verification, and token acquisition. */
const auth = new AuthModule(
    /* These values are required for our application to exhcange and get access to the resource data */
    /* client_id */ 'eb343a99-ba17-4cc1-87f5-86a9bbb64943',
    /* client_secret */ '',

    /* This information tells our server where to download the signing keys to validate the JWT that we received,
     * and where to get tokens. This is not configured for multi tenant; i.e., it is assumed that the source of the JWT and our application live
     * on the same tenant.
     */
    /* tenant */ 'common',
    /* stsDomain */ 'https://login.microsoftonline.com',
    /* discoveryURLsegment */ 'v2.0/.well-known/openid-configuration',
    /* tokenURLsegment */ '/oauth2/v2.0/token',

    /* Token is validated against the following values - NOTE: If you don't put a value, it won't validate against it */
    
    /* audience */ '', 
    /* scopes */ ['access_as_user'],
    /* issuer */ '',
);

/* A promisified express handler to catch errors easily */
const handler =
    (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
        (req, res, next) => callback(req, res, next)
            .catch(error => {
                /* If the headers are already sent then resort to the built in error handler */
                if (res.headersSent) {
                    return next(error);
                }

                /**
                 * If running development environment we send the error details back.
                 * Else we send the right code and message.
                 */
                if (env === 'development') {
                    return res.status(error.code || 500).json({ error });
                }
                else {
                    return res.status(error.code || 500).send(error.message);
                }
            });

/* Create the express app and add the required middleware */
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(morgan('dev'));
app.use(express.static('public'));
/* Turn off caching when debugging */
app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next()
});

/**
 * If running on development env, then use the locally available certificates.
 */
if (env === 'development') {
    const cert = {
        key: fs.readFileSync(path.resolve('./dist/certs/server.key')),
        cert: fs.readFileSync(path.resolve('./dist/certs/server.crt'))
    };
    https.createServer(cert, app).listen(44308, () => console.log('Server running on 44308'));
}
else {
    /**
     * We don't use https as we are assuming the production environment would be on Azure.
     * Here IIS_NODE will handle https requests and pass them along to the node http module
     */
    app.listen(process.env.port || 44308, () => console.log(`Server listening on port ${process.env.port}`));
}

app.get('/api/GraphToken', handler(async (req, res) => {
    /**
     * Only initializes the auth the first time
     * and uses the downloaded keys information subsequently.
     */
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });

    // 1. We don't pass a resource parameter because the token endpoint is Azure AD V2.
    // 2. Always ask for the minimal permissions that the application needs..
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All', 'Mail.Read',  'Calendars.Read']);

    return res.json(graphToken);
}));

/**
 * HTTP GET: /index.html
 * Loads the add-in home page.
 */
app.get('/index.html', handler(async (req, res) => {
    return res.sendfile('index.html');
}));


