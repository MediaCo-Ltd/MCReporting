//
// Functions for HS Reporting export Select Page
//

function cleardate(ctr) 
{
    document.getElementById(ctr).value = '';
}

function ShowDatePicker(Id,CboId,hId) 
{

    if (document.getElementById("cboSource").selectedIndex == 0) 
    {
        // do nothing
    }
    else 
    {
        document.getElementById(CboId).selectedIndex = 0;
        document.getElementById(hId).value = 0;
        //NewCal(Id, 'ddmmyyyy', false, 24);
        NewCssCal(Id, 'ddmmyyyy', 'arrow');
    }

}

function JSDateToDateSerial(inDate) 
{
    // Split the sent Uk date into 3 elements
    var mydate; 
    mydate = inDate.split("/");
    inDate = '';
    // Rearrange so its mmddyyy
    var swapdate;
    swapdate = mydate[1] + '/' + mydate[0] + '/' + mydate[2];
    // Date serial functions
    mydate = '';
    var d = new Date(swapdate);
    swapdate = '';
    var returnDateTime = 25569.0 + ((d.getTime() - (d.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
    d = '';
    return returnDateTime.toString().substr(0, 5);

    //document.getElementById("DateSerial1").value = returnDateTime.toString().substr(0, 5);

}
 
function ErrChk() 
{
    var emsg = document.getElementById("ErrBox").value
    if (document.getElementById("ErrBox").value.length > 1)
    { window.alert(emsg); }
}

function LogOff()
{
    window.location.replace("LogOff.asp");
}

function GoBackHSOption()
{
    window.location.replace('HsSelectOption.asp');
}

function GoBackHSCustom() 
{
    window.location.replace('HsCustomSelect.asp')
}

function ResetPage(DoWhat)
{
    if (DoWhat == '1') 
    {   // Only used for select export page so that return query string is removed
        window.location.replace('HsCustomSelect.asp');
    }
    else 
    {
        window.location.reload();
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

function ValidateExport() 
{
    
    // Only need to check any selected Start & End dates
    var chkStartDate = document.getElementById("hStartDate").value
    var chkEndDate = document.getElementById("hEndDate").value

    if (Number(chkStartDate) > 0 && Number(chkEndDate) == 0) 
    {
        alert('Invalid End date');
        return false;
    }

    if (Number(chkStartDate) == 0 && Number(chkEndDate) > 0) 
    {
        alert('Invalid Start date');
        return false;
    }

    if (Number(chkStartDate) > Number(chkEndDate)) 
    {
        alert('End date is before start date');
        return false;
    }
    else 
    {
        return true;
    }
}

function Relocate(DoWhat) 
{
    var result = ValidateExport();
    var LocationUrl = '';
    if (DoWhat == 'v') 
    {
        // relocate to display page where excel can be downloaded
        if (result == true) 
        {
            LocationUrl = 'HSCustomDisplay.asp'
            LocationUrl += '?lt=' + document.getElementById("hLogType").value;
            LocationUrl += '&re=' + document.getElementById("hReasonId").value;
            LocationUrl += '&us=' + document.getElementById("hUserId").value;
            LocationUrl += '&sh=' + document.getElementById("hShift").value;
            LocationUrl += '&sd=' + document.getElementById("hStartDate").value;
            LocationUrl += '&ed=' + document.getElementById("hEndDate").value;
            LocationUrl += '&ds=' + document.getElementById("hSource").value;
            LocationUrl += '&lc=' + document.getElementById("hLocation").value;

            window.location.replace(LocationUrl);
        }
    }        
}

function lockSubmit() 
{
    if (document.getElementById("hLogType").value != '') 
    {
        document.getElementById("btnView").disabled = false;
        document.getElementById("cboReasons").disabled = false;
        document.getElementById("cboLogUser").disabled = false;
        document.getElementById("cboShift").disabled = false;
        document.getElementById("cboSource").disabled = false;
        document.getElementById("cboStart").disabled = false;
        document.getElementById("cboEnd").disabled = false;
        document.getElementById("cboLocation").disabled = false;

        // document.getElementById("btnExport").disabled = false;
    }
    else 
    {
        document.getElementById("btnView").disabled = true;
        document.getElementById("cboReasons").disabled = true;
        document.getElementById("cboLogUser").disabled = true;
        document.getElementById("cboShift").disabled = true;
        document.getElementById("cboSource").disabled = true;
        document.getElementById("cboStart").disabled = true;
        document.getElementById("cboEnd").disabled = true;
        document.getElementById("cboLocation").disabled = true;

        //document.getElementById("btnExport").disabled = true;
    }
}

function UnlockSubmit() 
{
    document.getElementById("btnView").disabled = false;
    document.getElementById("cboReasons").disabled = false;
    document.getElementById("cboLogUser").disabled = false;
    document.getElementById("cboShift").disabled = false;
    document.getElementById("cboSource").disabled = false;
    document.getElementById("cboStart").disabled = false;
    document.getElementById("cboEnd").disabled = false;
    document.getElementById("cboLocation").disabled = false;

    //document.getElementById("btnExport").disabled = false;
}

function HSSelectLogType() 
{
    var strLogType = document.getElementById("cboLogType").value;

    // Any change to Type reset hidden fields
    document.getElementById("hReasonId").value = '0';
    document.getElementById("hUserId").value = '0';
    document.getElementById("hShift").value = '0';  
    
    if (document.getElementById("cboLogType").value != '') 
    {
        document.getElementById("hLogType").value = strLogType;        

        // SentType Resets Reasons
        getReasons(Number(strLogType));
        // SentType, SentReason Resets Users
        getUser(Number(strLogType), Number(0));
        // SentType, SentReason SentUser Resets Shift
        getShift(Number(strLogType), Number(0), Number(0));
        // SentType, Sending Page
        getLocation(Number(strLogType), 'Custom');
        UnlockSubmit();
    }
    else 
    {
        document.getElementById("hLogType").value = '';
        
        // Reset Reasons to All
        getReasons(Number(4));
        // Reset Users to All
        getUser(Number(4), Number(0));
        // Reset Shift to All
        getShift(Number(4), Number(0), Number(0));
        // SentType, Sending Page
        getLocation(Number(4), 'Custom');
        lockSubmit();
    }
}


function HSSelectReasons() 
{
    var strReason = document.getElementById("cboReasons").value;
    var LogType = document.getElementById("hLogType").value;

    // Any change to Reason reset hidden fields
    document.getElementById("hUserId").value = '0';
    document.getElementById("hShift").value = '0'; 

    if (document.getElementById("cboReasons").value != '') 
    {
        document.getElementById("hReasonId").value = strReason
        // SentType, SentReason Resets Users
        getUser(Number(LogType), Number(strReason));
        // SentType, SentReason SentUser Resets Shift
        getShift(Number(LogType), Number(strReason), Number(0));
    }
    else 
    {
        document.getElementById("hReasonId").value = '0'

        // Reset Users to All
        getUser(Number(LogType), Number(0));
        // Reset Shift to All
        getShift(Number(4), Number(0), Number(0));
    }
}

function HSSelectLogUser() 
{
    var strLogUser = document.getElementById("cboLogUser").value;
    var strSendType = document.getElementById("hLogType").value;
    var strSendReason = document.getElementById("hReasonId").value;

    // Any change to User reset hidden fields
    document.getElementById("hShift").value = '0'; 

    if (document.getElementById("cboLogUser").value != '') 
    {
        document.getElementById("hUserId").value = strLogUser
        // SentType, SentReason SentUser Resets Shift
        getShift(Number(strSendType), Number(strSendReason), Number(strLogUser));
    }
    else 
    {
        document.getElementById("hUserId").value = '0'
        // Reset Shift to All
        getShift(Number(4), Number(0), Number(0));
    }
}

function HSSelectShift() 
{

    var strShift = document.getElementById("cboShift").value;

    if (document.getElementById("cboShift").value != '') 
    {
        document.getElementById("hShift").value = strShift
    }
    else 
    {
        document.getElementById("hShift").value = '0'
    }

}


function HSSelectLocation() 
{
    var LocationId = document.getElementById("cboLocation").value;
    if (document.getElementById("cboLocation").value != '') 
    {
        document.getElementById("hLocation").value = LocationId;
    }
    else 
    {
        document.getElementById("hLocation").value = '0';
    }
}

// created date resolved date
function HSDateSource() 
{
    var strDateSource = document.getElementById("cboSource").value;
    if (document.getElementById("cboSource").value != '') 
    {
        document.getElementById("cboStart").innerHTML = '<option value="">Select Date</option>';
        //document.getElementById("cboStart").innerHTML += '<option value="0" >All</option>';
        document.getElementById("cboEnd").innerHTML = '<option value="">Select Date</option>';
        //document.getElementById("cboEnd").innerHTML += '<option value="0" >All</option>';
        document.getElementById("txtStart").value = '';
        document.getElementById("txtEnd").value = '';
        // Populate dropdowns with dates based on source Ajax function
        //getDates(Number(strDateSource));

        document.getElementById("hStartDate").value = '0';
        document.getElementById("hEndDate").value = '0';
        document.getElementById("EndSame").disabled = false;
        document.getElementById("hSource").value = strDateSource;
    }
    else 
    {
        document.getElementById("cboStart").innerHTML = '<option value="">Select Date</option>';
        document.getElementById("cboEnd").innerHTML = '<option value="">Select Date</option>';
        document.getElementById("hStartDate").value = '0'
        document.getElementById("hEndDate").value = '0';
        document.getElementById("hSource").value = '0';
        document.getElementById("txtStart").value = '';
        document.getElementById("txtEnd").value = '';
        document.getElementById("EndSame").disabled = true;
    }
}

function HSStartDate() 
{
    var strStartDate = document.getElementById("cboStart").value;

    if (document.getElementById("cboStart").value != '') 
    {
        document.getElementById("hStartDate").value = strStartDate;
        document.getElementById("txtStart").value = '';
    }
    else 
    {
        if (document.getElementById("txtStart").value == '')
        { document.getElementById("hStartDate").value = '0'; }
    }
}

function HSEndDate() 
{
    var strEndDate = document.getElementById("cboEnd").value;

    if (document.getElementById("cboEnd").value != '') 
    {
        document.getElementById("hEndDate").value = strEndDate;
        document.getElementById("txtEnd").value = '';
    }
    else 
    {
        if (document.getElementById("txtEnd").value == '')
        { document.getElementById("hEndDate").value = '0'; }

    }
}

function HSEndSame() 
{
    if (document.getElementById("cboStart").value != '') 
    {
        document.getElementById("hEndDate").value = document.getElementById("hStartDate").value;
        document.getElementById("cboEnd").selectedIndex = document.getElementById("cboStart").selectedIndex;
        document.getElementById("txtStart").value = '';
        document.getElementById("txtEnd").value = ''; 
    }
    else 
    {
        document.getElementById("txtEnd").value = document.getElementById("txtStart").value;
        document.getElementById("hEndDate").value = document.getElementById("hStartDate").value;
        document.getElementById("cboEnd").selectedIndex = document.getElementById("cboStart").selectedIndex;
    }
}
