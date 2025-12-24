
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

var getReasons = function(SentIds) 
{
    document.getElementById("cboReason").innerHTML = '';
    var ReasonStr = "xSentIds=" + SentIds;
    
    //alert(SentIds)

    doAJAXCall('NcAjaxGetReasons.asp', 'POST', ReasonStr, showReasonsResponse);
}

var showReasonsResponse = function(oXML) 
{
    // get the response text, into a variable
    var ReasonResponse = oXML.responseText;

    //document.getElementById("txtDetails").innerHTML = ReasonResponse;
    if (ReasonResponse == "") 
    {
        document.getElementById("cboReason").innerHTML = "<option value='' >Select Issue/Problem</option>";
        //document.getElementById("cboReason").innerHTML += "<option value='102' >Other</option>";
        alert("No matching default reasons for this group/work centre!");
    }
    else 
    {
        document.getElementById("cboReason").innerHTML = ReasonResponse;
        //document.getElementById("cboReason").innerHTML += "<option value='102' >Other</option>";
    }
}


