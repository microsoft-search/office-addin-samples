var text = "";
var debug = false;
var graphUrl = "https://graph.microsoft.com";
var version = "beta";
var ssoToken = "";
var graphToken = "";
var size = 25;
var from = 0;

function writeDebug(text) {
    if (debug) {        
        document.getElementById("debug").innerHTML += text + "<br/>";
    }
}

$(document).ready(function () {

    $("#btnSearch").click(function () {

        var query = $("#tbQuery").val();
        var entityType = $("#entityType").val();

        size = $("#tbCount").val();
        from = $("#tbSkip").val();

        writeDebug("Searching for " + query);

        doSearch(query, entityType, from, size);
    });

    $("#cbDebug").click(function () {

        var divDebug = $("#debug");
        debug = $("#cbDebug").is(':checked');
        
        if (debug) {
            divDebug.css('display', 'block');
        }
        else {
            divDebug.css('display', 'none');
        }
        
    });

    getGraphToken();

});

function getGraphToken() {

    $.ajax({
        type: "POST",
        url: "https://localhost:44308/api/Graph/Token",
        data: {  },
        success: function (data) {
            graphToken = data;
        },
        error: function (error) {
            writeDebug(error);
        }
    });

}

function doSearch(query, entityType, from, size) {

    var request = {};
    request.entityType = "microsoft.graph." + entityType;
    request.query = {};
    request.query["query_string"] = {};
    request.query["query_string"].query = query;
    request.from = from;
    request.size = size;
    request["_sources"] = [];
    request["_sources"].push("from");
    request["_sources"].push("to");
    request["_sources"].push("subject");
    request["_sources"].push("body");

    var jsonObj = {};
    var requests = [];
    requests.push(request);
    jsonObj["requests"] = requests;

    var jsonString = JSON.stringify(jsonObj);
    var url = graphUrl + "/" + version + "/search";
    
    document.getElementById("request").innerHTML = "<h2>Request</h2>" + jsonString;

    $.ajax({
        type: "POST",
        url: url,
        headers: {
            "Authorization": "Bearer " + graphToken
        },
        data: jsonString,
        contentType: "application/json",
        success: function (data) {

            var jsonString = JSON.stringify(data);
            document.getElementById("response").innerHTML = "<h2>Request</h2>" + jsonString;

            var results = $("#results");
            results.css('display', 'block');
            results.empty();
            var html = "";
            html += "<ul>";

            //loop through results
            data.value[0].hitsContainers[0].hits.forEach(function (item) {
                var link = "";
                if (item._source.webLink) {
                    link = item._source.webLink;
                }
                if (item._source.webUrl) {
                    link = item._source.webUrl;
                }
                if (entityType == 'event') {
                    link = "https://outlook.office365.com/calendar/view/month";
                }

                if (entityType == 'driveItem') {
                    item._summary = item._source.name;
                }

                html += "<li><a href='" + link + "'>" + item._summary + "</a></li>";
            });

            html += "</ul>";
            results.append(html);
        },
        error: function (error) {

            var results = $("#results");
            results.empty();

            var res = $("#response");
            res.empty();

            if (debug)
                document.getElementById("response").innerHTML = "<h2>Response</h2>" + JSON.stringify(error);

            writeDebug("Error getting results");
        }
    }).done(function (data) {

    }).fail(function (error) {

        var results = $("#results");
        results.empty();

        var res = $("#response");
        res.empty();

        if (debug) {
            document.getElementById("response").innerHTML = "<h2>Response</h2>" + JSON.stringify(error);

            writeDebug("Error getting results");
        }
            
    });

}