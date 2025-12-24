
// Ajax Functions For Machine Faults 

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

var getFaultGroup = function(SentGroups) 
{
    document.getElementById("cboFaultGroup").innerHTML = ''
    var PostStr = "xMg=" + SentGroups;
    
    doAJAXCall('MfAjaxGetFaultGroup.asp', 'POST', PostStr, showStatusResponse);
}

var showStatusResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    //document.getElementById("txtError").innerHTML = StatusResponse;

    if (StatusResponse == "") 
    {
        document.getElementById("cboFaultGroup").innerHTML = "<option value='' >Select&nbsp;Fault&nbsp;Group</option>"

//        document.getElementById("chkRecurring").checked = false;
//        document.getElementById("chkRecurring").disabled = true;
//        document.getElementById("cboFaultGroup").innerHTML = "";
//        document.getElementById("GroupLabel").style.visibility = "hidden";
//        document.getElementById("GroupCbo").style.visibility = "hidden";       
        
        alert("No matching groups for this machine type!");
    }
    else 
    {
        document.getElementById("cboFaultGroup").innerHTML = StatusResponse;
    }
}


//SetDormant

var SetMfDormant = function(RecordId) 
{

    var PostStatusStr = "xRecordId=" + RecordId + "&xStatus=1";

    doAJAXCall('MfAjaxSetRecordStatus.asp', 'POST', PostStatusStr, showMfRecordStatusResponse);
}

var UnSetMfDormant = function(RecordId) {

    var PostStatusStr = "xRecordId=" + RecordId + "&xStatus=0";

    doAJAXCall('MfAjaxSetRecordStatus.asp', 'POST', PostStatusStr, showMfRecordStatusResponse);
}

var showMfRecordStatusResponse = function(oXML) {
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;


    if (StatusResponse == "") 
    {
        alert("Unable to change status\nOr session has expired");
        window.location.reload();
    }
    else {
        alert(StatusResponse);
        window.location.reload();
    }
}