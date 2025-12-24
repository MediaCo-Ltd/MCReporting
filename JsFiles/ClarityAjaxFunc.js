
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

var getJobData = function(ref) 
{
    var PostStr = "xQr=" + ref;
    doAJAXCall('AjaxGetjob.asp', 'POST', PostStr, showStatusResponse);
}

var showStatusResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    
    //document.getElementById("txtError").innerHTML = StatusResponse;

    if (StatusResponse == "") 
    {
        alert("Job doesn't exist, or is Outdoor");
    }
    else 
    {
        if (StatusResponse == Number(0)) 
        {
            alert("Account for this job\nIs on hold or not authorised. Contact Paul");
            window.location.reload();
        }

        else 
        {
            var DetailWinParams, DetailURL, DetailWindow, W, H, L;

            W = screen.availWidth - 200;
            H = screen.availHeight - 150;
            L = (screen.availWidth - W) / 2;

            DetailWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
            DetailWinParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
            DetailWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=40';
            DetailURL = 'ClarityShowDetail.asp?Job=' + (StatusResponse) + "&log=0";
            DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);
            // Reload window to clear box
            window.location.reload();
        }

    }
}

var SetRedo = function(Qid) 
{
    var PostStr = "xQid=" + Qid;
    doAJAXCall('AjaxSetRedo.asp', 'POST', PostStr, showRedoResponse);    
}

var showRedoResponse = function(oXML) {
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    //document.getElementById("txtError").value = StatusResponse;

    if (StatusResponse == "") 
    {
        alert("Unable to set redo status");
        window.location.reload();
    }
    else 
    {
        if (StatusResponse == "01")
        { alert("Unable to set redo status"); }        
        window.location.reload();
    }
}

var getJobLin = function(BomJobId) 
{

    if (document.getElementById('Lin' + BomJobId).getAttribute("title") == '') 
    {
        var PostStr = "xBomId=" + BomJobId;
        doAJAXCall('AjaxGetLinear.asp', 'POST', PostStr, showLinearResponse);
    }    
}

var showLinearResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    var SplitData = StatusResponse.split("#");
    

    if (SplitData[0] == "") 
    {
        // Do Nothing alert("No data");
    }
    else 
    {
        document.getElementById('Lin' + SplitData[1]).setAttribute("title", SplitData[0]);
    }
}

var getJobLinItems = function(Id) 
{
    var Splitdata = Id.split("#");
    if (document.getElementById('Lin' + Splitdata[1]).getAttribute("title") == '') 
    {
        var PostStr = "xBomId=" + Id;
        doAJAXCall('AjaxGetLinearItems.asp', 'POST', PostStr, showLinearItemResponse);
    }
}

var showLinearItemResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    //alert(StatusResponse);

    var SplitData = StatusResponse.split("#");


    if (SplitData[0] == "") 
    {
        // Do Nothing alert("No data");
    }
    else 
    {
        document.getElementById('Lin' + SplitData[1]).setAttribute("title", SplitData[0]);
    }
}

var getJobLinSubstrate = function(SubId) 
{

    if (document.getElementById('Lin' + SubId).getAttribute("title") == '') 
    {
        var PostStr = "xBomId=" + SubId + '#' + document.getElementById('hSubstrate').value;

        doAJAXCall('AjaxGetLinearSubstrate.asp', 'POST', PostStr, showLinearSubstrateResponse);
    }
}

var showLinearSubstrateResponse = function(oXML) 
{
    // get the response text, into a variable
    var StatusResponse = oXML.responseText;

    //alert(StatusResponse);

    var SplitData = StatusResponse.split("#");


    if (SplitData[0] == "") 
    {
        // Do Nothing alert("No data");
    }
    else 
    {
        document.getElementById('Lin' + SplitData[1]).setAttribute("title", SplitData[0]);
    }
}
