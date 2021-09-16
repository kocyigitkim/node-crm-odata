import { HttpsAgent } from "agentkeepalive";

export async function crm_ntlm_auth(agent : HttpsAgent, url: string, username: string, password: string, domain: string, hostname: string, requestCallback?: Function) : Promise<any> {
    if (!requestCallback) requestCallback = ((v) => v);

    return new Promise(async (resolve, reject) => {

        var ntlm = require('httpntlm').ntlm;
        var async = require('async');
        var httpreq = require('httpreq');
        var keepaliveAgent = agent;

        var options = {
            url: url,
            username: username,
            password: password,
            workstation: hostname,
            domain: domain
        };

        async.waterfall([
            function (callback) {
                var type1msg = ntlm.createType1Message(options);
                var req = requestCallback({
                    headers: {
                        'Connection': 'keep-alive',
                        'Authorization': type1msg
                    },
                    agent: keepaliveAgent
                });
                httpreq[(req.method || "get")](options.url, req, callback);
            },

            function (res, callback) {
                if (!res.headers['www-authenticate'])
                    return callback(new Error('www-authenticate not found on response of second request'));

                var type2msg = ntlm.parseType2Message(res.headers['www-authenticate']);
                var type3msg = ntlm.createType3Message(type2msg, options);

                setImmediate(function () {
                    var req = requestCallback({
                        headers: {
                            'Connection': 'Close',
                            'Authorization': type3msg
                        },
                        allowRedirects: false,
                        agent: keepaliveAgent
                    });
                    httpreq[(req.method || "get")](options.url, req, callback);
                });
            }
        ], function (err, res) {
            if (err) return reject(err);
            resolve(res);
        });

    });
}