// Module of KeyPay API Calls
// Set API Key first with Keypay.setApiKey("APIKey")
var Keypay = (function() {
  var apiKey = false;
  var businessId = false;

  // Fetches data from Keypay
  // urlType determines base of URL ("bus" for .../business/businessId/, "ess" for .../ess/)
  function _apiGet(urlType, apiUrl, apiData) {

    if (!apiKey) {
      throw new Error("Error, please set Keypay.apiKey before making API call");
    }
    
    var encodedApiKey = Utilities.base64EncodeWebSafe(apiKey);

    if (urlType == "bus") {
      if (!businessId) {
        throw new Error("Error, please set Keypay.businessId before making API call");
      }
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/business/' + businessId;

    } else if (urlType == "ess") {
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/ess/';

    } else if (urlType == "noId") {
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/business/';

    } else {
      throw new Error("apiGet called with invalid URL type: " + urlType);
    }

    var headers = {
          'Authorization': 'Basic ' + encodedApiKey
        };
  
    var options = {
      'method': 'get',
      'contentType': 'application/json',
      'headers': headers,
      'muteHttpExceptions' : true
    };
    
    var concatUrl = urlBase + apiUrl;
    
    // If there is data in apiData, join and concatenate to URL
    if (typeof apiData !== "undefined") {
      concatUrl = concatUrl + '?' + _convertObjToUrl(apiData);
    }

    var response = _tryCall(concatUrl, options);
    var responseCode = response.getResponseCode();

    if (responseCode == 200) {
      var responseData = JSON.parse(response);
      return responseData;
    } else {
      Logger.log("Uh oh, " + concatUrl + " sent response code " + responseCode);
      return false;
    }

  }  


  // Posts data to Keypay
  // urlType determines base of URL ("bus" for .../business/businessId/, "ess" for .../ess/)
  function _apiPost(urlType, apiUrl, body, apiData) {

    if (!apiKey) {
      throw new Error("Error, please set Keypay.apiKey before making API call");
    }
    
    var encodedApiKey = Utilities.base64EncodeWebSafe(apiKey);

    if (urlType == "bus") {
      if (!businessId) {
        throw new Error("Error, please set Keypay.businessId before making API call");
      }
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/business/' + businessId;

    } else if (urlType == "ess") {
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/ess/';

    } else {
      throw new Error("apiGet called with invalid URL type: " + urlType);
    }

    var headers = {
          'Authorization': 'Basic ' + encodedApiKey
        };
  
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'headers': headers,
      'muteHttpExceptions' : true,
      'payload': JSON.stringify(body)
    };
    
    var concatUrl = urlBase + apiUrl;
    
    // If there is data in apiData, join and concatenate to URL
    if (typeof apiData !== "undefined") {
      concatUrl = concatUrl + '?' + _convertObjToUrl(apiData);
    }

    var response = _tryCall(concatUrl, options);
    var responseCode = response.getResponseCode();

    if (responseCode == 201) {
      return response;
    } else {
      Logger.log("Uh oh, " + concatUrl + " sent response code " + responseCode);
      Logger.log("Content = " + response.getContentText());
      return response;
    }

  }  


  // Updates data in Keypay
  // urlType determines base of URL ("bus" for .../business/businessId/, "ess" for .../ess/)
  function _apiPut(urlType, apiUrl, body, apiData) {

    if (!apiKey) {
      throw new Error("Error, please set Keypay.apiKey before making API call");
    }
    
    var encodedApiKey = Utilities.base64EncodeWebSafe(apiKey);

    if (urlType == "bus") {
      if (!businessId) {
        throw new Error("Error, please set Keypay.businessId before making API call");
      }
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/business/' + businessId;

    } else if (urlType == "ess") {
      var urlBase = 'https://api.yourpayroll.com.au/api/v2/ess/';

    } else {
      throw new Error("apiGet called with invalid URL type: " + urlType);
    }

    var headers = {
          'Authorization': 'Basic ' + encodedApiKey
        };
  
    var options = {
      'method': 'put',
      'contentType': 'application/json',
      'headers': headers,
      'muteHttpExceptions' : true,
      'payload': JSON.stringify(body)
    };
    
    var concatUrl = urlBase + apiUrl;
    
    // If there is data in apiData, join and concatenate to URL
    if (typeof apiData !== "undefined") {
      concatUrl = concatUrl + '?' + _convertObjToUrl(apiData);
    }

    var response = _tryCall(concatUrl, options);
    var responseCode = response.getResponseCode();

    if (responseCode == 200) {
      return response;
    } else {
      Logger.log("Uh oh, " + concatUrl + " sent response code " + responseCode);
      Logger.log("Content = " + response.getContentText());
      return response;
    }

  }  

  // Tries to call url up to 5 times. Returns data on 200 response code, false on anything else
  function _tryCall(url, options) {
    Logger.log("Calling " + url);
    var maxAttempts = 5;
    var attempts = 0;

    while (attempts < maxAttempts) {
      try {

        var response = UrlFetchApp.fetch(url, options);
        // Logger.log(url + " returned " + response.getResponseCode());
        return response;

      } catch (e) {

        attempts++;
        Logger.log("Attempt failed, " + attempts + " failures so far");

      }
    }
    
    Logger.log("Max API attempts reached, giving up");
    throw new Error("API call failed " + attempts + " times: url = " + url);
  }


  function _convertObjToUrl(obj) {
    var str = "";
    for (var key in obj) {
      if (str != "") {
        str += "&";
      }
      str += (key == 'filter' ? '$' + key : key) + "=" + encodeURI(obj[key]);
    }
    return str;
  }

  // --------------------- Public Methods Below ------------------------------
  function initialise() {
    var docProperties = PropertiesService.getDocumentProperties();
    var ui = SpreadsheetApp.getUi();
    if (docProperties.getProperty('API_KEY') == null) {
      ui.alert("Please set your API key first");
      return false;
    }
    if (docProperties.getProperty('BUSINESS_ID') == null) {
      ui.alert("Please set your Business ID first");
      return false;
    }

    Keypay.setApiKey(docProperties.getProperty('API_KEY'));
    Keypay.setBusinessId(docProperties.getProperty('BUSINESS_ID'));
    return true;
  }


  // Sets apiKey
  function setApiKey(input) {
    apiKey = input;
  }

  // Sets businessId
  function setBusinessId(input) {
    businessId = input;
  }

  // Lists all the earnings lines for a pay run.
  function listEarningsLines(payrunId) {
    return _apiGet('bus', '/payrun/' + payrunId + '/earningslines');
  }

  // This endpoint returns a list of employees. The details are a subset of the 'unstructured' employee endpoint.
  function listEmployees() {
    return _apiGet('bus', '/employee/details');
  }

  // Returns the unsttuctured data for all employees on a given payschedule (or all if not provided)
  function listEmployeesByPayschedule(payScheduleId) {
    var emps = [];
    var page = 0;
    var pageStr = "?$top=1";
    var paySchedStr = payScheduleId == undefined ? "" : "&filter.payScheduleId=" + payScheduleId;
    var str = pageStr + paySchedStr;
    var response = _apiGet('bus', '/employee/unstructured' + str);
    while (response && response.length != 0) {
      pageStr = "?$top=100&$skip=" + 100 * page;
      var str = pageStr + paySchedStr;
      response = _apiGet("bus", "/employee/unstructured" + str);
      emps = emps.concat(response);
      page++;
    }
    return emps;
  }

  // Lists the details for all of the documents in the business.
  function listBusinessDocuments() {
    return _apiGet('bus', '/document');
  }

  // Lists all documents visible to this employee, including both business and employee documents.
  function listEssDocuments(employeeId) {
    return _apiGet('ess', employeeId + '/document');
  }

  
  // Get a list of pay runs associated with the business.
  function listPayRuns() {
    return _apiGet('bus', '/payrun');
  }

  
  // Gets the pay run with the specified ID.
  function getPayRun(payRunId) {
    return _apiGet('bus', '/payrun/' + payRunId);
  }


  // Grants a user access to the employee.
  function grantEmployeeAccess(employeeId, email, name) {
    var body = {
      "email": email,
      "name": name,
      "suppressNotificationEmails": true
    };

    return _apiPost('bus', '/employee/' + employeeId + '/access', body);
  }


  // Creates a business location.
  function createLocation(locationBody) {
    return _apiPost('bus', '/location', locationBody);
  }


  // Updates a business location.
  function updateLocation(locationBody) {
    return _apiPut('bus', '/location/' + locationBody.id, locationBody);
  }


  // Lists all the locations for a business.
  function listBusinessLocations() {
    return _apiGet('bus', '/location');
  }


  // Retrieves the details of the business with the specified ID.
  function getBusinessDetails() {
    return _apiGet('bus', '');
  }


  // Lists all pay rate templates for a business.
  function listPayRateTemplates() {
    return _apiGet('bus', '/payratetemplate');
  }


  // Creates a new pay rate template for the business.
  function createPayRateTemplate(prtBody) {
    return _apiPost('bus', '/payratetemplate', prtBody);
  }


  // Updates the pay rate template with the specified ID.
  function updatePayRateTemplate(prtBody) {
    return _apiPut('bus', '/payratetemplate/' + prtBody.id, prtBody);
  }


  // Lists all the pay categories for the business.
  function listPayCategories() {
    return _apiGet('bus', '/paycategory');
  }


  // Lists all businesses for the API Key
  function listBusinesses() {
    return _apiGet('noId', '');
  }


  // Gets the pay rates for this employee.
  function getPayRates(empId) {
    return _apiGet('bus', '/employee/' + empId + '/payrate');
  }


  // Lists all the pay schedules for the business.
  function getPaySchedules() {
    return _apiGet('bus', '/payschedule');
  }


  // Generates a timesheet report.
  function getTimesheetReport(fromDate, toDate, payScheduleId) {
    // request.fromDate=2019-05-15T00:00:00&request.toDate=2019-05-15T00:00:00&request.statuses=Processed&request.payScheduleId={{payScheduleId}}
    var str = "?request.status=AnyExceptRejected&request.fromDate=" + fromDate + "&request.toDate=" + toDate;
    if (payScheduleId != undefined) {
      str += "&request.payScheduleId=" + payScheduleId;
    }
    return _apiGet('bus', '/report/timesheet' + str);
  }

  function listEmploymentAgreements() {
    return _apiGet('bus', '/employmentagreement');
  }

  function getEmploymentAgreementbyId(id) {
    return _apiGet('bus', '/employmentagreement/' + id);
  }

  function createLeaveCategory(body) {
    return _apiPost('bus', '/leavecategory', body);
  }

  function updatePayCategory(body) {
    return _apiPut('bus', '/paycategory/' + body.id, body);
  }

  function createPayCategory(body) {
    return _apiPush('bus', '/paycategory/', body);
  }

  function getLeaveCategories() {
    return _apiGet('bus', '/leavecategory');
  }


  return {
    start: initialise,
    setApiKey: setApiKey,
    setBusinessId: setBusinessId,
    getEarningsForPayrun: listEarningsLines,
    getEmployees: listEmployees,
    getBusinessDocuments: listBusinessDocuments,
    getEssDocuments: listEssDocuments,
    getPayRuns: listPayRuns,
    getPayRun: getPayRun,
    grantEmployeeAccess: grantEmployeeAccess,
    getEmployeesByPayschedule: listEmployeesByPayschedule,
    createLocation: createLocation,
    updateLocation: updateLocation,
    getLocations: listBusinessLocations,
    getBusinessDetails: getBusinessDetails,
    getPayRateTemplates: listPayRateTemplates,
    createPayRateTemplate: createPayRateTemplate,
    updatePayRateTemplate: updatePayRateTemplate,
    getPayCategories: listPayCategories,
    getBusinesses: listBusinesses,
    getPayRates: getPayRates,
    getPaySchedules: getPaySchedules,
    getTimesheetReport: getTimesheetReport,
    getEAs: listEmploymentAgreements,
    getEAbyID: getEmploymentAgreementbyId,
    createLeaveCategory: createLeaveCategory,
    updatePayCategory: updatePayCategory,
    getLeaveCategories: getLeaveCategories,
    createPayCategory: createPayCategory
  }

})();


function setApiKey() {
  var docProperties = PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Enter API Key");
  if (response.getSelectedButton() == ui.Button.OK) {
    docProperties.setProperty('API_KEY', response.getResponseText());
  } else {
    Logger.log("Selected cancel");
    return false;
  }
  return true;
}


function setBusinessId() {
  var docProperties = PropertiesService.getDocumentProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Enter Business ID");
  if (response.getSelectedButton() == ui.Button.OK) {
      docProperties.setProperty('BUSINESS_ID', response.getResponseText());
      Keypay.setApiKey(docProperties.getProperty('API_KEY'));
      Keypay.setBusinessId(docProperties.getProperty('BUSINESS_ID'));
      var business = Keypay.getBusinessDetails();
      ss.toast("Business set to " + business.name);        
  } else {
      Logger.log("Selected cancel");
      return false;
  }
  return true;
}