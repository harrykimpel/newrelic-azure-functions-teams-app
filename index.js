const https = require('https');

module.exports = function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    //const name = (req.query.name || (req.body && req.body.name));
    //const responseMessage = name
    //    ? "Hello, " + name + ". This HTTP triggered function executed successfully."
    //    : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

    var newRelicUserKey = GetEnvironmentVariable("NEWRELIC_USER_KEY");
    //context.log(newRelicUserKey);
    var newRelicDashboardEntityGuid = GetEnvironmentVariable("NEWRELIC_DASHBOARD_ENTITYGUID");
    //context.log(newRelicDashboardEntityGuid);
    var refreshInterval = GetEnvironmentVariable("REFRESH_INTERVAL");
    //context.log(refreshInterval);
    var isRefreshPage = GetEnvironmentVariable("IS_REFRESH_PAGE");
    //context.log(isRefreshPage);

    var optionsNerdGraph = {
        host: 'api.newrelic.com',
        port: 443,
        path: '/graphql',
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'API-Key': newRelicUserKey
        }
    };

    var graphQuery =
    {
        "query": "mutation {dashboardCreateSnapshotUrl(guid: \"" + newRelicDashboardEntityGuid + "\")}"
    };

    var refreshPageCode = '';
    if (isRefreshPage == 'true') {
        refreshPageCode = `setTimeout(onRefresh, ` + refreshInterval + `);`;
    }

    var content = `<!DOCTYPE html>
                <html>
                    <head>
                        <script src='https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js'></script>
                        <style>
                        .display-name {
                            font-family: "Segoe UI",system-ui,"Apple Color Emoji","Segoe UI Emoji",sans-serif;
                        }
                        </style>
                    </head>
                    <body>
                        <span class="display-name">Microsoft Teams User: </span><span id="user" class="display-name"></span>
                        <div id="id01">#CHUNK#</div>
                        <script>
                                    microsoftTeams.initialize();
                            microsoftTeams.getContext((context) => {
                                let userId = document.getElementById('user');
                                userId.innerHTML = context.userPrincipalName;
                            });
                            function onRefresh() {
                                location.reload();
                            }
                            `+ refreshPageCode + `
                        </script>
                        #NEWRELIC#
                        <input type="button" name="Button" value="Refresh" onclick="onRefresh()">
                    </body>
                </html>`;
    var chunk = '';

    const myReq = https.request(optionsNerdGraph, function (res) {
        context.log('STATUS: ' + res.statusCode);
        res.setEncoding('utf8');
        res.on('data', (d) => {
            chunk += d;
        });

        res.on('end', (d) => {
            //content = content.replace('#CHUNK#', chunk);
            content = content.replace('#CHUNK#', '');

            var json = JSON.parse(chunk);
            var imgUrl = json.data.dashboardCreateSnapshotUrl;
            imgUrl = imgUrl.replace('?format=PDF', '?format=PNG');
            context.log('imgUrl: ' + imgUrl);

            var newrelicChart = `<img id="nrChart" src="` + imgUrl + `" />`;

            content = content.replace('#NEWRELIC#', newrelicChart);

            context.res = {
                status: 200,
                headers: {
                    'Content-Type': 'text/html'
                },
                body: content
            }
            context.done();
        })
    })

    myReq.on('error', function (e) {
        context.log('problem with request: ' + e.message);

        content = content.replace('#CHUNK#', 'An error occured: ' + e.message);

        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'text/html'
            },
            body: content
        }
        context.done();
    });

    //context.log('query: ' + JSON.stringify(graphQuery));
    myReq.write(JSON.stringify(graphQuery));

    myReq.end();
}

function GetEnvironmentVariable(name) {
    return process.env[name];
}
