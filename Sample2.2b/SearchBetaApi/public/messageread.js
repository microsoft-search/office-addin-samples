//https://docs.microsoft.com/en-us/office/dev/add-ins/reference/manifest/versionoverrides

//https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/AttachmentDemoWeb/MessageRead.js

(function () {
    "use strict";
  
      var text = "";
      var debug = true;
      var graphUrl = "https://graph.microsoft.com";
      var version = "beta";
      var ssoToken = "";
      var graphToken = "";
  
      function writeDebug(text) {
          if (debug) {
              var divDebug = $("#debug");
              divDebug.css('display', 'block');
              
              document.getElementById("debug").innerHTML += text + "<br/>";
          }
      }
  
      Office.onReady(function () {
  
          try {
              text += Office.context.requirements.isSetSupported("IdentityApi")
                  ? "IdentityApi is supported"
                  : "IdentityApi NOT supported";
          } catch (e) {
              text += "Error: " + e;
          }
  
          writeDebug(text);
  
          //set initial item textbox
          var item = Office.context.mailbox.item;
          var query = item.subject;
          $("#tbQuery").val(query);
  
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
  
                  writeDebug(text + "Auth called success: token : " + ssoToken);
  
                  //trade for graph token...
                  getGraphToken(ssoToken);
              }
              else {
  
                  writeDebug( "Auth called (error):" + result.error.code);
  
                  writeDebug( "Auth called (error) : " + JSON.stringify(result));
  
                  if (result.error.code === 13000) {
                      writeDebug( "This version of Office is not supported. Please upgrade.");
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
                      writeDebug( "Auth called: token : " + ssoToken);
                  } else {
                      // Handle error
                  }
  
                  if (result.error.code === 13003) {
                      
                  } else {
                      // Handle error
                  }
  
                  if (result.error.code === 13007) {
                      // SSO is not supported for domain user accounts, only
                      // work or school (Office 365) or Microsoft Account IDs.
                      // OR YOU HAVE FIDDLER ENABLED

                      writeDebug( "Auth called (error): Are you running Fiddler?");

                  } else {
                      // Handle error
                  }
              }
          });
      }
  
      function getGraphToken(token) {
  
          $.ajax({
              type: "GET",
              headers: {
                "Authorization": "Bearer " + ssoToken
                },
              url: "https://localhost:44308/api/GraphToken",
              //contentType: "application/json; charset=utf-8",
              success: function (data) {
                  graphToken = data;
              },
              error: function (error) {
                  writeDebug(error);
              }
          });
  
      }
  
      function doSearch(query, entityType) {
  
          var request = {};
          request.entityType = "microsoft.graph." + entityType;
          request.query = {};
          request.query["query_string"] = {};
          request.query["query_string"].query = query;
          request.from = 0;
          request.size = 25;
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
  
                  var results = $("#results");
                  results.css('display', 'block');
                  results.empty();
  
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
  
                      document.getElementById("results").innerHTML += "<a href='" + link + "'>" + item._summary + "</a><br>";
                  });
  
  
  
                  if (debug)
                      document.getElementById("response").innerHTML = "<h2>Response</h2>" + JSON.stringify(data);
              },
              error: function (error) {
  
                  var results = $("#results");
                  results.empty();
  
                  if (debug)
                      document.getElementById("results").innerHTML = "<h2>Response</h2>" + JSON.stringify(error);
              }
          }).done(function (data) {
              
          }).fail(function (error) {
  
              var results = $("#results");
              results.empty();
  
              if (debug)
                  document.getElementById("results").innerHTML = "<h2>Response</h2>" + JSON.stringify(error);
          });
  
      }
  
    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
    function buildAttachmentsString(attachments) {
      if (attachments && attachments.length > 0) {
        var returnString = "";
        
        for (var i = 0; i < attachments.length; i++) {
          if (i > 0) {
            returnString = returnString + "<br/>";
          }
          returnString = returnString + attachments[i].name;
        }
  
        return returnString;
      }
  
      return "None";
    }
  
    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
      return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }
  
    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
      if (addresses && addresses.length > 0) {
        var returnString = "";
  
        for (var i = 0; i < addresses.length; i++) {
          if (i > 0) {
            returnString = returnString + "<br/>";
          }
          returnString = returnString + buildEmailAddressString(addresses[i]);
        }
  
        return returnString;
      }
  
      return "None";
      }
  
      function getCurrentItem(accessToken) {
          // Get the item's REST ID
          var itemId = getItemRestId();
  
          // Construct the REST URL to the current item
          // Details for formatting the URL can be found at
          // /previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-a-message-rest
          var getMessageUrl = Office.context.mailbox.restUrl +
              '/v2.0/me/messages/' + itemId;
  
          $.ajax({
              url: getMessageUrl,
              dataType: 'json',
              headers: { 'Authorization': 'Bearer ' + accessToken }
          }).done(function (item) {
              // Message is passed in `item`
              var subject = item.Subject;
        }).fail(function (error) {
                      // Handle error
                  });
      }
  
    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;
        $('#subject').text(item.subject);
  
      $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
      $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
      $('#itemClass').text(item.itemClass);
      $('#itemId').text(item.itemId);
      $('#itemType').text(item.itemType);
  
      $('#message-props').show();
  
      $('#attachments').html(buildAttachmentsString(item.attachments));
      $('#cc').html(buildEmailAddressesString(item.cc));
      $('#conversationId').text(item.conversationId);
      $('#from').html(buildEmailAddressString(item.from));
      $('#internetMessageId').text(item.internetMessageId);
      $('#normalizedSubject').text(item.normalizedSubject);
      $('#sender').html(buildEmailAddressString(item.sender));
      $('#subject').text(item.subject);
      $('#to').html(buildEmailAddressesString(item.to));
    }
  
    // Helper function for displaying notifications
    function showNotification(header, content) {
      $("#notificationHeader").text(header);
      $("#notificationBody").text(content);
      messageBanner.showBanner();
      messageBanner.toggleExpansion();
    }
    
  })();