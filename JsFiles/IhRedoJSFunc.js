//
//
// OH Redo Functions


function LogOff() 
{
    window.location.replace("LogOff.asp");
}

function ViewData() 
{
    window.location.replace("IhrSelectDate.asp");
}

function LoadRedo(RedoId) 
{
    var ReturnPage = document.getElementById("ReturnAddress").value
    window.location.replace("IhrShowData.asp?Rid=" + RedoId + '&rp=' + ReturnPage);
}

function TimeTest() 
{
    window.location.replace("IhrDisabled.asp");
}

function AddWc(CboId) 
{
    if (document.getElementById("CboDept" + [CboId]).value != '') 
    {
        var e = document.getElementById("CboDept" + [CboId]);
        var WcName = e.options[e.selectedIndex].value;

        var Gname = e.options[e.selectedIndex].text;
        document.getElementById("Grp" + [CboId]).value = Gname;

        // Reset Code text
        document.getElementById("ReasonTxt" + [CboId]).value = '';  
        
        document.getElementById("WCID" + [CboId]).value = WcName;
       
        //Get Qc group id for passing to ajax to populate reason codes
        var GrpData = WcName.split("#");

        //alert(GrpData[0] + ' ' + GrpData[1] + ' ' + GrpData[2])
        var GrpRow = CboId;
        GroupReason(GrpRow, GrpData[1], GrpData[2]);

//        if (GrpData[0] == '7') 
//        {
//            document.getElementById("Dept" + [CboId]).value = '3';
//        }
        // for testing
        //document.getElementById("Reason" + [CboId]).value = GrpRow + '#' + GrpData[1];        
        //document.getElementById("Qty" + [CboId]).value = document.getElementById("OrigQty" + [CboId]).value;
    }
    else 
    {
        document.getElementById("Grp" + [CboId]).value = '';
        document.getElementById("WCID" + [CboId]).value = '';
        document.getElementById("Reason" + [CboId]).value = '';
        document.getElementById("CboReason" + [CboId]).style.visibility = "hidden"
        document.getElementById("CboReason" + [CboId]).innerHTML = '<option value="" >Select&nbsp;Reason</option>';
        //document.getElementById("Qty" + [CboId]).value = document.getElementById("OrigQty" + [CboId]).value;
        document.getElementById("Reason" + [CboId]).disabled = true;
        document.getElementById("Code" + [CboId]).value = '';
        document.getElementById("ReasonTxt" + [CboId]).value = '';
        
    }
}

function AddCodes(CodeRowId) 
{
    if (document.getElementById("CboReason" + [CodeRowId]).value != '') 
    {
        var e = document.getElementById("CboReason" + [CodeRowId]);
        var CodeNo = e.options[e.selectedIndex].value;
        var ReasonTxt = e.options[e.selectedIndex].text;

        document.getElementById("ReasonTxt" + [CodeRowId]).value = ReasonTxt;
        document.getElementById("Code" + [CodeRowId]).value = CodeNo;
        document.getElementById("Reason" + [CodeRowId]).disabled = false;
               
    }
    else 
    {
        document.getElementById("ReasonTxt" + [CodeRowId]).value = '';
        document.getElementById("Code" + [CodeRowId]).value = '';
        document.getElementById("Reason" + [CodeRowId]).value = '';
        document.getElementById("Reason" + [CodeRowId]).disabled = true;
    }
}

function ShowAllRows() 
{
    var TotalRows = Number(document.getElementById("TotalRows").value);
    var Rcount = 0;
    var NoRows = 0;
    //TotalRows -= 1;
        do 
        {
            if (document.getElementById("chk" + [Rcount]).disabled == true) 
            {
                NoRows +=2;
                Rcount++; 
            }
            else 
            {
                document.getElementById("Row" + [Rcount]).style.visibility = "visible";
                document.getElementById("chk" + [Rcount]).checked = true;
                Rcount++;
            }
        }
        while (Rcount < TotalRows);

        if(TotalRows - NoRows < 0)
        {
            document.getElementById("ActualRows").value = 0
        }
        else
        {
        document.getElementById("ActualRows").value = TotalRows - NoRows;
        }
        document.getElementById("TickAll").style.visibility = "hidden";
        
}

function ShowRow(RowId) 
{
    var ActRow = Number(document.getElementById("ActualRows").value);
    
    if (document.getElementById("Row" + [RowId]).style.visibility == "collapse") 
    {
        document.getElementById("Row" + [RowId]).style.visibility = "visible";
        ActRow += 1;
                
    }
    else 
    {
        document.getElementById("Row" + [RowId]).style.visibility = "collapse";
        document.getElementById("WCID" + [RowId]).value = '';
        //document.getElementById("Qty" + [RowId]).value = document.getElementById("OrigQty" + [RowId]).value;
        document.getElementById("Qty" + [RowId]).value = '0'
        document.getElementById("CboDept" + [RowId]).selectedIndex = 0;
        document.getElementById("CboReason" + [RowId]).innerHTML = '<option value="" >Select&nbsp;Reason</option>';
        document.getElementById("CboReason" + [RowId]).style.visibility = "hidden"
        document.getElementById("Code" + [RowId]).value = '';
        ActRow -=1
    }

    document.getElementById("ActualRows").value = ActRow;
    if (ActRow == 0)
    { document.getElementById("btnSubmit").disabled = true; }

    if (ActRow == 0) 
    {
        if (Number(document.getElementById("TotalRows").value) > 1) 
        {
            document.getElementById("TickAll").style.visibility = "visible";
        }

    } 
    

}

function EnableSubmit() 
{
    if (document.getElementById("btnSubmit").disabled == true) 
    {
        document.getElementById("btnSubmit").disabled = false;
    }   
}

function DisableEnterKey(e) 
{
    var key;

    if (window.event)
        key = window.event.keyCode;     //IE
    else
        key = e.which;     //firefox

    if (key == 13)
    { return false; }
    else
    { return true; }
}


function ValidateRedoJobNo() 
{

    var redonumber = document.getElementById("txtJobNo").value;
    var ReplaceRef = redonumber.replace("REF", "");

    if (isNaN(ReplaceRef) || ReplaceRef == '' || ReplaceRef == '0')    
    {
        alert("Enter a valid number");
        return false;
    }
	else 
    {
        document.body.style.cursor = 'wait';
        document.getElementById('loading').innerHTML = "Loading data, please wait...";
        document.getElementById("subsection1").style.opacity = 0;
        document.getElementById("subsection2").style.opacity = 0;
        document.getElementById("logoffL").style.color = "#ffffff";
    }

}

function ValidateData() 
{
    var QtyError = 0;
    var QtyText = '';
    var i = 0;
    var RowTotal = Number(document.getElementById("TotalRows").value);
    var ActualRows = Number(document.getElementById("ActualRows").value);
    var ReasonError = 0;
    var ReasonErrorText = '';
    var NothingSelected = 0;
    var SpError = 0;
    var SpText = '';
	

    
    RowTotal -= 1;
        do {
            if (Number(document.getElementById("Code" + [i]).value) > 0) 
            {
                if (Number(document.getElementById("Qty" + [i]).value) > Number(document.getElementById("OrigQty" + [i]).value)) 
                {
                    QtyError += 1;
                    QtyText += 'Redo qty for Item ' + document.getElementById("ItemIdAlpha" + [i]).value + ' is greater than original\n'
                }

                if (Number(document.getElementById("Qty" + [i]).value) == Number(0) )
                {
                    QtyError += 1;
                    QtyText += 'Redo qty for Item ' + document.getElementById("ItemIdAlpha" + [i]).value + ' is zero\n'
                }
            }
            else 
            {
                // do nothing
            }
     
            if (document.getElementById("WCID" + [i]).value != '') 
            {
                if (Number(document.getElementById("Code" + [i]).value) > 0) 
                {
                    // Check if other selected 205, 311, 409, 505, 602
                    if (Number(document.getElementById("Code" + [i]).value) == Number(311) || Number(document.getElementById("Code" + [i]).value) == Number(205) || Number(document.getElementById("Code" + [i]).value) == Number(409) || Number(document.getElementById("Code" + [i]).value) == Number(505) || Number(document.getElementById("Code" + [i]).value) == Number(602))  
                    {
                        if (document.getElementById("Reason" + [i]).value == '')
                        {                        
                        ReasonError += 1;
                        ReasonErrorText += 'Enter additional info for Item ' + document.getElementById("ItemIdAlpha" + [i]).value + '\n'
                        }                    
                    }
                }
                else 
                {
                    ReasonError += 1;
                    ReasonErrorText += 'Select a reason for Item ' + document.getElementById("ItemIdAlpha" + [i]).value + '\n'
                }
            }

            if (document.getElementById("SpId" + [i]).value == '' && Number(document.getElementById("Code" + [i]).value) > 0) 
            {
                if (document.getElementById("SpTrue" + [i]).value == 'True' && Number(document.getElementById("Qty" + [i]).value) < Number(document.getElementById("OrigQty" + [i]).value))  
                {
                    SpError += 1;
                    SpText += 'No SP Items for Item ' + document.getElementById("ItemIdAlpha" + [i]).value + '\n'
                }
            }
            else 
            {
                // do nothing
            }
            
            i++;
        }
        while (i <= RowTotal);

        if (QtyError > 0) 
        {
            alert(QtyText);
            return false;
        }
        else
        {

            if (ReasonError > 0) 
            {
                alert(ReasonErrorText);
                return false;
            }
            else 
            {
                if (SpError > 0) 
                {
                    alert(SpText);
                    return false;
                }
                else
                { 
					document.getElementById("btnSubmit").style.visibility = "hidden";
                    document.getElementById("reset").style.visibility = "hidden";
					return true; 
				}
            }
        }
}



function DataPageLoad() 
{
    if (document.getElementById("TotalRows").value == 1)
        document.getElementById("SelectTxt").innerHTML = '';
}

function PageLoadChk() 
{
    var focusHere = document.frmJobNo.txtJobNo;
    focusHere = focusHere.focus();
    return Echk();
}


function Echk() 
{
    var emsg = document.getElementById("ErrBox").value
    if (document.getElementById("ErrBox").value.length > 1) 
    {
        if (emsg == 'Search Error') 
        {
            emsg = 'Search has returned multiple results \nEnter more data to refine search'
        }
        window.alert(emsg);
    }
    document.getElementById("ErrBox").value = ''
}

//
// Ajax Functions
// 

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
    else 
    {
        // Not support by this browser, alert the user
        alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");
    }
}

var GroupReason = function(GrpRowId, Grp, WcId ) 
{
    document.getElementById("Reason" + [GrpRowId]).value = '';

    document.getElementById("Dept" + [GrpRowId]).value = Grp
    var PostStr = "xGrpId=" + Grp;
    PostStr += "&xRowId=" + GrpRowId;
    PostStr += "&xWcId=" + WcId;
    doAJAXCall('IhrAjaxGetReason.asp', 'POST', PostStr, showReasonResponse);
}

var showReasonResponse = function(oXML) {
    // get the response text, into a variable
    var ReasonResponse = oXML.responseText;
    var SplitData = ReasonResponse.split("#");

    document.getElementById("ReasonTest").value = ReasonResponse

    if (ReasonResponse == "BOF") {
        //alert("No matching Reasons for this group!\nAdd manual reason");
        // document.getElementById("ReasonTest").value = ReasonResponse
        //alert("System Error");

    }
    else {
        document.getElementById("CboReason" + SplitData[0]).style.visibility = "visible"
        document.getElementById("CboReason" + SplitData[0]).innerHTML = '';
        document.getElementById("CboReason" + SplitData[0]).innerHTML = SplitData[1];
        document.getElementById("Code" + SplitData[0]).value = SplitData[2]
        document.getElementById("btnSubmit").disabled = false;
        if (SplitData[2] != '')
        { document.getElementById("Reason" + SplitData[0]).disabled = false; }
        

    }
}