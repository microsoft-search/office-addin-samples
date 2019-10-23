'use strict';

(function () {

    var text = "";
    var debug = true;
    var graphUrl = "https://graph.microsoft.com";
    var version = "beta";
    var ssoToken = "";
    var graphToken = "";
    var limit = 25;
    var from = 0;

    function writeDebug(text) {
        if (debug) {
            var divDebug = $("#debug");
            divDebug.css('display', 'block');

            document.getElementById("debug").innerHTML += text + "<br/>";
        }
    }

    function checkSets() {
        try {
            text = Office.context.requirements.isSetSupported("ExcelApi")
                ? "ExcelApi is supported<br/>"
                : "ExcelApi NOT supported<br/>";
        } catch (e) {
            text = "Error: " + e;
        }

        try {
            text += Office.context.requirements.isSetSupported("IdentityApi")
                ? "IdentityApi is supported<br/>"
                : "IdentityApi NOT supported<br/>";
        } catch (e) {
            text += "Error: " + e;
        }

        writeDebug(text);
    }

    Office.onReady(function () {

        OfficeExtension.config.extendedErrorLogging = true;

        checkSets();

        $("#btnSearch").click(function () {
            var query = $("#tbQuery").val();
            var entityType = $("#entityType").val();
            doSearch(query, entityType);
        });

        var options = { forceConsent: true, forceAddAccount: false };
        options = { forceConsent: false, forceAddAccount: false };
        login(options);
    });

    function login(options) {
        Office.context.auth.getAccessTokenAsync(options, function (result) {
            if (result.status === "succeeded") {

                //we got an identity token...
                ssoToken = result.value;

                writeDebug ("Auth called success: token : " + ssoToken );

                //trade for graph token...
                getGraphToken(ssoToken);
            }
            else {
                writeDebug("Auth called (error):" + result.error.code );

                if (result.error.code === 13000) {
                    writeDebug("This version of Office is not supported. Please upgrade.");
                } else {
                    // Handle error
                }

                if (result.error.code === 13001) {
                    var options = { forceConsent: true };
                    login(options);
                } else {
                    // Handle error
                }

                if (result.error.code === 13002) {
                    //show the consent link...
                    writeDebug(text + "Auth called: token : " + ssoToken);
                } else {
                    // Handle error
                }

                if (result.error.code === 13003) {
                    writeDebug("Check your Office account, it must be an O365 or Microsoft Account");
                } else {
                    // Handle error
                }

                if (result.error.code === 13007) {
                    writeDebug("Check your Office account, it must be an O365 or Microsoft Account");
                } else {
                    // Handle error
                }
            }
        });
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showNotification('The selected text is:', '"' + result.value + '"');
                    } else {
                        showNotification('Error:', result.error.message);
                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    };

    function getGraphToken(token) {

        $.ajax({
            type: "POST",
            url: "https://localhost:44308/api/Graph/Token",
            //contentType: "application/json; charset=utf-8",
            data: { userToken: ssoToken },
            success: function (data) {
                graphToken = data;
            },
            error: function (error) {
                writeDebug(error);
            }
        });
    }

    function intToCell(count) {

        var cell = "";

        var loops = count / 65;
        loops = Math.ceil(loops);

        for (var i = 0; i < loops; i++) {

            if (count > 65) {
                cell += "A";
            }
            else {
                var res = String.fromCharCode(count+65).toUpperCase();
                cell += res;
            }
        }

        return cell;
    }

    //http://buildingofficeaddins.com/sync/
    //https://gist.github.com/ATrogolo/938ccce2d29448b349c3ec989d53ea8e
    var excelHeaders = [];
    var excelData = [];
    var tableRange = "B1:E1";

    function parseResult(jsonData) {
        var data = jsonData.value[0].hitsContainers[0].hits[0];

        var substrateHeaders = Object.keys(data);
        var itemHeaders = Object.keys(data._source);

        excelHeaders = [];
        excelHeaders = excelHeaders.concat(substrateHeaders);
        excelHeaders = excelHeaders.concat(itemHeaders);

        var cell = intToCell(excelHeaders.length);
        tableRange = "B1:" + cell + "1";

        excelData = [];
        var count = 0;

        jsonData.value[0].hitsContainers[0].hits.forEach(function (item) {

            count = 0;

            var newItem = [];

            excelHeaders.forEach(function (key) {
                var val = "";

                if (substrateHeaders.indexOf(key)!=-1) {
                    val = item[key];
                }
                else {
                    val = item._source[key];
                }

                //if its empty...append empty value
                if (val === null) {
                    newItem.push("");
                }
                else {
                    var type = typeof val;

                    try {
                        console.log(key + ":" + type + ":" + val.length + ":" + val);
                    }
                    catch (e) {
                        console.log(val);
                    }

                    if (type == "object") {
                        val = JSON.stringify(val);
                        //val = "";
                    }

                    newItem.push(val);
                }
            });

            excelData.push(newItem);
        });

        console.log(count + " properties.");
    }

    function addRows(tableName, jsonData) {

        Excel.run(function (context) {

            parseResult(jsonData);
            
            var currentWorksheet = context.workbook.worksheets.getItemOrNullObject("Sheet1");

            let excelTable = currentWorksheet.tables.getItemOrNullObject(tableName);

            return context.sync().then(function () {

                //delete the table each time...
                if (!excelTable.isNullObject) {
                    excelTable.delete();
                }

                excelTable = currentWorksheet.tables.add(tableRange, true /* hasHeaders */);
                excelTable.name = tableName;
                excelTable.getHeaderRowRange().values = [excelHeaders];

                excelTable.rows.add(0, excelData);                
            });

        }).catch(function (error) {

            console.log('error: ' + error);

            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                writeDebug(JSON.stringify(error.debugInfo));
            };

        });
    }

    function doSearch(query, entityType, from) {

        var request = {};
        request["entityTypes"] = [];
        request["entityTypes"].push("microsoft.graph." + entityType);

        request.query = {};
        request.query["query_string"] = {};
        request.query["query_string"].query = query;
        request.from = from;
        request.size = limit;

        request["stored_fields"] = [];
        request["stored_fields"].push("from");
        request["stored_fields"].push("to");
        request["stored_fields"].push("subject");
        request["stored_fields"].push("body");

        var jsonObj = {};
        var requests = [];
        requests.push(request);
        jsonObj["requests"] = requests;

        var jsonString = JSON.stringify(jsonObj);
        var url = graphUrl + "/" + version + "/search/query";

        writeDebug(jsonString);

        $.ajax({
            type: "POST",
            url: url,
            headers: {
                "Authorization": "Bearer " + graphToken
            }, 
            data: jsonString,
            contentType: "application/json",
            success: function (data) {

                addRows("Results", data);

                if (debug)
                    document.getElementById("response").innerHTML = "<h2>Response</h2>" + JSON.stringify(data);
            },
            error: function (error) {
                document.getElementById("results").innerHTML = "<h2>Response</h2>" + JSON.stringify(error);
            }
        }).done(function (data) {


        }).fail(function (error) {

            if (debug)
                writeDebug(JSON.stringify(error));

        });

    }
})();