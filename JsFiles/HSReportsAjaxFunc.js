
// Ajax Functions 

function XHConn() {
    var xmlhttp, bComplete = false;
    try { xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); }
    catch (e) {
        try { xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); }
        catch (e) {
            try { xmlhttp = new XMLHttpRequest(); }
            catch (e) { xmlhttp = false; } 
        } 
    }
    if (!xmlhttp) return null;
    this.connect = function(sURL, sMethod, sVars, fnDone) {
        if (!xmlhttp) return false;
        bComplete = false;
        sMethod = sMethod.toUpperCase();
        try {
            if (sMethod == "GET") {
                //open(method, url + query string, async)
                xmlhttp.open(sMethod, sURL + "?" + sVars, true);
                sVars = "";
            }
            else {
                //open(method, url, async)
                xmlhttp.open(sMethod, sURL, true);
                xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
            }
            xmlhttp.onreadystatechange = function() {
                if (xmlhttp.readyState == 4 && !bComplete) {
                    bComplete = true;
                    fnDone(xmlhttp);
                } 
            };
            xmlhttp.send(sVars);
        }
        catch (z) { return false; }
        return true;
    };
    return this;
}

// doAJAXCall : Generic AJAX Handler, used with XHConn
// Author : Bryce Christensen (www.esonica.com)
// PageURL : the server side page we are calling
// ReqType : either POST or GET, typically POST
// PostStr : parameter passed in a query string format 'param1=foo&param2=bar'
// FunctionName : the JS function that will handle the response

var doAJAXCall = function(PageURL, ReqType, PostStr, FunctionName) 
{

    // create the new object for doing the XMLHTTP Request
    var myConn = new XHConn();

    // check if the browser supports it
    if (myConn) {

        // XMLHTTPRequest is supported by the browser, continue with the request

        myConn.connect('' + PageURL + '', '' + ReqType + '', '' + PostStr + '', FunctionName);
    }
    else {
        // Not support by this browser, alert the user
        alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");
    }
}


//####################################################

var getReasons = function(SentType) 
{
    document.getElementById("cboReasons").innerHTML = '';
    var ReasonStr = "xSentType=" + SentType;

    doAJAXCall('HsAjaxGetReasons.asp', 'POST', ReasonStr, showReasonsResponse);
}

var showReasonsResponse = function(oXML) 
{
    // get the response text, into a variable
    var ReasonResponse = oXML.responseText;

    //document.getElementById("txtDetails").innerHTML = ReasonResponse;
    if (ReasonResponse == "") 
    {
        document.getElementById("cboReasons").innerHTML = "<option value='' >Select&nbsp;Reason</option>";
        alert("No matching reasons for this log type!");
    }
    else 
    {
        document.getElementById("cboReasons").innerHTML = ReasonResponse;
    }
}

var getUser = function(SentType, SentReason) 
{
    document.getElementById("cboLogUser").innerHTML = '';
    var PostHsUserStr = "xType=" + SentType + "&xReason=" + SentReason;


    //alert(PostUserStr)
    doAJAXCall('HsAjaxGetUsers.asp', 'POST', PostHsUserStr, showUserResponse);
}

var showUserResponse = function(oXML) 
{
    // get the response text, into a variable
    var UserResponse = oXML.responseText;

    //document.getElementById("txtDetails").innerHTML = UserResponse;
    if (UserResponse == "")
    {
        document.getElementById("cboLogUser").innerHTML = "<option value='' >Select&nbsp;User</option>";
    }
    else {
        document.getElementById("cboLogUser").innerHTML = UserResponse;
    }
}


var getShift = function(SentType, SentReason, SentUser) 
{
    document.getElementById("cboShift").innerHTML = '';
    var PostShiftStr = "xType=" + SentType + "&xReason=" + SentReason + "&xUser=" + SentUser;

    doAJAXCall('HsAjaxGetShift.asp', 'POST', PostShiftStr, showShiftResponse);
}

var showShiftResponse = function(oXML) {
    // get the response text, into a variable
    var ShiftResponse = oXML.responseText;

    //document.getElementById("txtDetails").innerHTML = ShiftResponse;
    if (ShiftResponse == "") 
    {
        document.getElementById("cboShift").innerHTML = "<option value='' >Select&nbsp;Shift</option>";
    }
    else 
    {
        document.getElementById("cboShift").innerHTML = ShiftResponse;
    }
}

var getLocation = function(ExcludedList,ReturnType) 
{
    document.getElementById("cboLocation").innerHTML = '';
    var PostLocationStr = "xExcluded=" + ExcludedList + "&xType=" + ReturnType;

    doAJAXCall('HsAjaxGetLocation.asp', 'POST', PostLocationStr, showLocationResponse);
}

var showLocationResponse = function(oXML) 
{
    // get the response text, into a variable
    var LocationResponse = oXML.responseText;

    // document.getElementById("txtDetails").innerHTML = LocationResponse;
    
    if (LocationResponse == "") 
    {
        document.getElementById("cboLocation").innerHTML = "<option value='' >Select&nbsp;Location</option>";
    }
    else 
    {
        document.getElementById("cboLocation").innerHTML = LocationResponse;
    }
}

// Riddor Status

var SetHsRiddor = function(RecordId) 
{

    var PostStatusStr = "xRecordId=" + RecordId ;

    doAJAXCall('HsAjaxSetRiddorStatus.asp', 'POST', PostStatusStr, showHsRiddorStatusResponse);
}

var showHsRiddorStatusResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;


    if (StatusResponse == "") 
    {
        alert("Unable to change status\nOr session has expired");
        window.location.reload();
    }
    else 
    {
        alert(StatusResponse);
        window.location.reload();
    }
}



//SetDormant

var SetHsDormant = function(RecordId) 
{

    var PostStatusStr = "xRecordId=" + RecordId + "&xStatus=1";

    doAJAXCall('HsAjaxSetRecordStatus.asp', 'POST', PostStatusStr, showHsRecordStatusResponse);
}

var UnSetHsDormant = function(RecordId) 
{

    var PostStatusStr = "xRecordId=" + RecordId + "&xStatus=0";

    doAJAXCall('HsAjaxSetRecordStatus.asp', 'POST', PostStatusStr, showHsRecordStatusResponse);
}

var showHsRecordStatusResponse = function(oXML) {
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;


    if (StatusResponse == "") 
    {
        alert("Unable to change status\nOr session has expired");
        window.location.reload();
    }
    else 
    {
        alert(StatusResponse);
        window.location.reload();
    }
}



