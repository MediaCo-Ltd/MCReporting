//
// functions for HS Reporting
//
//################## Common



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
    window.location.replace('HSSelectOption.asp');
}

function ShowOrder(e) 
{
    var upkey;

    if (window.event)
    { upkey = window.event.keyCode; }     //IE
    else
    { upkey = e.which; }    //firefox

    if (upkey == 13) 
    {
        var ordertoget = document.getElementById("JumpTxt").value;
        if (isNaN(ordertoget) || ordertoget == '' || ordertoget == '0')
        //if (rowtoget != '')
        {
            alert('Enter a valid number')
        }
        else {
            window.location.replace("HSEditLog.asp?id=" + ordertoget + "&sp=0");
        }
    }
}

function DisableEnterKey(e,objId) 
{
    var key;

    if (window.event)
        key = window.event.keyCode;     //IE
    else
        key = e.which;     //firefox

    if (key == 13) 
    {
        return false;
    }
    
}

function EnableSubmit(EnForm) 
{
    if (document.getElementById("btnSubmit").disabled == true) 
    {
        document.getElementById("btnSubmit").disabled = false;
    }

    if (EnForm == 'LoginPw')
    { document.getElementById("txtPassword").value = document.getElementById("txtUserChar").value; }
}



function ResetPage() 
{
    window.location.reload();
}

function TabToNext(e, objId) 
{
    var upkey;

    if (window.event)
        upkey = window.event.keyCode;     //IE
    else
        upkey = e.which;     //firefox

    if (upkey == 13) 
    {
        if (objId == 'JumpTxt') 
        {
            var rowtoget = document.getElementById(objId).value;
            document.getElementById("JumpTxt").value = '';
            if (rowtoget != '') 
            {
                document.getElementById(rowtoget).focus();
            }
        }
        else
        { document.getElementById(objId).focus(); }
    }
}

//############################# Option Page

function HSOptionsLoad() 
{
    // Firefox doesn't reset if return from back button
    // ok once added no-store to Response.CacheControl = "no-cache"

    if (document.getElementById("Locked").value == "True") 
    {
        var MsgText
        MsgText = "System Admin has locked this site for maintenance";
        document.getElementById("logoffL").style.visibility = "hidden";
        document.getElementById("logoffR").style.visibility = "hidden";
        document.getElementById("Logo").style.opacity = 1;
        //document.getElementById("Data").style.opacity = .3;
        document.getElementById('Options').style.opacity = .3;
        //document.getElementById('Footer').style.opacity = 1;
        alert(MsgText);
        window.location.replace("SystemLocked.html");
        return;
    }
    else {
        document.body.style.cursor = 'default';
        //document.getElementById('loading').innerHTML = '';
        //init();
        return ErrChk();
    }
}

function RelocateOption(strUrl) 
{
    window.location.replace(strUrl);
}

//########################### Any view Page

function ViewRecordsLoad() 
{
    if (document.getElementById("Locked").value == "True") 
    {
        var MsgText
        MsgText = "System Admin has locked this site for maintenance";
        alert(MsgText);
        window.location.replace("SystemLocked.html");
    }
}

function HSAddLogLoad()
{
    if (document.getElementById("Locked").value == "True") 
    {
        var MsgText
        MsgText = "System Admin has locked this site for maintenance";
        alert(MsgText);
        window.location.replace("SystemLocked.html");
    }
    else 
    {
        document.getElementById("txtDetails").focus();
    }
}

function HSAddReason() 
{
    var ReasonId = document.getElementById("cboReason").value;
    var SplitData = ReasonId.split("#");
    
    if (document.getElementById("cboReason").value != '')
    {
        //document.getElementById("hReasonId").value = ReasonId
        document.getElementById("hReasonId").value = SplitData[0];    
        document.getElementById("cboSeverity").disabled = false;
        document.getElementById("cboSeverity").selectedIndex = 0;
        document.getElementById("hSeverity").value = '';
        document.getElementById("cboLocation").disabled = false;
        document.getElementById("cboLocation").selectedIndex = 0;
        //document.getElementById("chkResolved").disabled = false;
        //document.getElementById("chkResolved").checked = false;

        if (SplitData[0] == '14') 
		{
            // index is 1 more than value not working??
            // This works, but writing inner html better
            // document.getElementById("cboLocation").selectedIndex = 22;
            document.getElementById("cboLocation").innerHTML = "<option value='24' >Mezz Gate</option>";
            document.getElementById("hLocation").value = '24';
            document.getElementById("cboLocation").disabled = true;            
        }
        else
        {
            // Reset location
            document.getElementById("hLocation").value = ''
            getLocation(SplitData[1],'Log');
        }
             
    }
    else 
    {
        document.getElementById("chkResolved").checked = false;
        //document.getElementById("chkResolved").disabled = true;
        document.getElementById("cboSeverity").selectedIndex = 0;
        document.getElementById("cboSeverity").disabled = true;
        document.getElementById("cboLocation").selectedIndex = 0;
        document.getElementById("cboLocation").disabled = true;      
        document.getElementById("hSeverity").value = '';
        document.getElementById("hReasonId").value = '';
        document.getElementById("hGroupId").value = '';
        document.getElementById("hLocation").value = '';
    }
}

function HSSeverity() 
{
    var SeverityId = document.getElementById("cboSeverity").value;
    if (document.getElementById("cboSeverity").value != '') 
    {
        document.getElementById("hSeverity").value = SeverityId;
    }
    else 
    {
        document.getElementById("hSeverity").value = '';
    }
}

function HSAddLocation() 
{
    var LocationId = document.getElementById("cboLocation").value;
    if (document.getElementById("cboLocation").value != '') 
    {
        document.getElementById("hLocation").value = LocationId;
    }
    else 
    {
        document.getElementById("hLocation").value = '';
    }
}

function HSAddRiddor() 
{
    var RiddorId = document.getElementById("cboRiddor").value;
    var RiddorSplit = RiddorId.split("#");
    var TodaySerial = document.getElementById("hTodayDateSerial").value;
    var CreatedSerial = document.getElementById("hLogCreatedSerial").value;
    var DayDiff = TodaySerial - CreatedSerial;
    var DaysLeft = 0;

    if (document.getElementById("cboRiddor").value != '') 
    {
        document.getElementById("hRiddor").value = RiddorSplit[0];
        document.getElementById("hRiddorDays").value = RiddorSplit[1];

        if (DayDiff == 0)
        {DaysLeft = Number(RiddorSplit[1]); }
        else
        { DaysLeft = Number(RiddorSplit[1]) - Number(DayDiff) ; }

        if (document.getElementById("frmName").value == "frmAddLog") 
        {
            document.getElementById("RiddorDays").innerHTML = 'You have ' + DaysLeft + ' Days to submit';
        }

        if (document.getElementById("frmName").value == "frmEditLog") 
        {
            if (DayDiff > Number(RiddorSplit[1])) 
            {
                document.getElementById("RiddorDays").innerHTML = 'You have missed the to submit period';
            }
            else
            {
                if (document.getElementById("cboRiddor").value == 0) 
                {
                    document.getElementById("RiddorDays").innerHTML = '';
                    document.getElementById("hRiddorDays").value = 0;
                }

                else
                { document.getElementById("RiddorDays").innerHTML = 'You have ' + DaysLeft + ' Days to submit'; }
            }
            
        }
        
        //alert('TodaySerial = ' + TodaySerial + '\nCreatedSerial = ' + CreatedSerial + '\nDayDiff = ' + DayDiff + '\nDaysLeft = ' + DaysLeft);
        
       
    }
    else 
    {
        document.getElementById("hRiddor").value = 0;
        document.getElementById("hRiddorDays").value = 0;
    }
}


function HSEditLogLoad() 
{

    if (document.getElementById("Locked").value == "True") 
    {
        var MsgText
        MsgText = "System Admin has locked this site for maintenance";
        alert(MsgText);
        window.location.replace("SystemLocked.html");
    }
    else 
    {
        var RecordLocked = document.getElementById("hLockStatus").value;
        var LockedByName = document.getElementById("LockedByName").value;
        
        if (RecordLocked == 'True') 
        {
            document.getElementById("btnSubmit").style.visibility = "hidden";
            alert('Record is currently locked.\n' + LockedByName + ' is editing it')
        }

        if (document.getElementById("txtResponse").value == '')
        { document.getElementById("txtResponse").focus(); }
        else
        { document.getElementById("txtResponseNew").focus(); }
            
        //getFaultGroup(document.getElementById("hType").value);
    }
}


function HSLoadEditOrder(EditId) 
{
    var Sender = document.getElementById("Sendpage").value;

    window.location.replace("HSEditLog.asp?id=" + EditId + "&sp=" + Sender);    

}


function HSLoadViewOrderUser(EditId) 
{

    window.location.replace("HsViewUserLog.asp?id=" + EditId );  

}

function HSValidateLog(SendingPage)
{
    var ValidateMsg = ''
    var ValidateError = 0

    if (SendingPage == 'Add') 
    {
        if (document.getElementById("txtDetails").value == '') 
        {
            ValidateMsg += "Enter some details\n";
            ValidateError += 1
        }

        if (document.getElementById("hReasonId").value == '') 
        {
            ValidateMsg += "Select a reason\n";
            ValidateError += 1
        }

        if (document.getElementById("hSeverity").value == '') 
        {
            ValidateMsg += "Select severity\n";
            ValidateError += 1
        }

        if (document.getElementById("hLocation").value == '') 
        {
            ValidateMsg += "Select location\n";
            ValidateError += 1
        }
               
    }
    else 
    {
        if (document.getElementById("hHasNotes").value == 'False') 
        {
            if (document.getElementById("txtResponse").value == '') 
            {
                ValidateMsg += "Enter some comments\n";
                ValidateError += 1
            }        
        }       
        
        
    }

    if (ValidateError > 0) 
    {
        alert("Please correct the following...\n\n" + ValidateMsg);
        return false;
    }
    else 
    {        
        return true; 
    }
}


//############################ Admin Functions ###################################

function HSReasonNameChanged() 
{
    var NameChange = Number(document.getElementById("hReasonChange").value)
    NameChange += 1;
    document.getElementById("hReasonChange").value = NameChange;

}

function HSReasonSelect() 
{
    var strReasonId = document.getElementById("cboReason").value;
    if (document.getElementById("cboReason").value != '') 
    {
        document.getElementById("btnSubmit").disabled = false;
        document.getElementById("hReasonId").value = strReasonId;
    }
    else 
    {
        document.getElementById("hReasonId").value = '';
        document.getElementById("btnSubmit").disabled = true;
    }

}

function HSReasonTypeSelect() 
{
    var strReasonTypeId = document.getElementById("cboReasonType").value;
    if (document.getElementById("cboReasonType").value != '') 
    {
        document.getElementById("btnSubmit").disabled = false;
        document.getElementById("hReasonType").value = strReasonTypeId;
    }
    else 
    {
        document.getElementById("hReasonType").value = '';
        document.getElementById("btnSubmit").disabled = true;
    }

}

function HSValidateReason(FormName) 
{

    var ReasonCount = 0
    var ReasonTxt = document.getElementById("txtReasonID").value.toUpperCase()
    var ReasonChk = document.getElementById("hReasonList").value.toUpperCase()
    var ReasonArr = ReasonChk.split("#");

    if (document.getElementById("hReasonChange").value == '0' && FormName == "Edit")
    { }  // do Nothing
    else
        for (ReasonCount = 0; ReasonCount < ReasonArr.length; ReasonCount++) 
        {
            if (ReasonTxt == ReasonArr[ReasonCount]) 
            {
                alert("Reason in use, enter a new one");
            return false;
        }
    }

    if (document.getElementById("txtReasonID").value == '') 
    {
        alert("Enter a Reason");
        document.getElementById("txtReasonID").focus();
        return false
    }

    if (document.getElementById("hReasonType").value == '') 
    {
        alert("Select a Reason Type");
        return false
    }

}

function AddImage(Mode) 
{

    var UploadWinParams, UploadURL, UploadWindow, W, H, L, T;

    H = 400;
    W = 500;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    UploadWinParams = 'toolbar=no,location=no,status=no,menubar=no,copyhistory=no';
    UploadWinParams += ',scrollbars=no,statusbar=no,resizable=no';
    UploadWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=' + T;

    if (Mode == 'E')
    { UploadURL = "HsUploadEdit.asp"; }
    else
    { UploadURL = "HsUploadAdd.asp"; }

    UploadWindow = window.open(UploadURL, '_blank', UploadWinParams);

}

function LoadImage(Id) 
{


    var ShowImagesParams, ShowImagesURL, ShowImagesWindow, W, H, L, T;
    ShowImagesURL = 'HsShowImages.asp?Id=' + (Id);

    W = 1000;
    H = 800;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    ShowImagesParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    ShowImagesParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    ShowImagesParams += ',width= 1000, height=800, left=' + L + ',top=' + T;

    ShowImagesWindow = window.open(ShowImagesURL, '_blank', ShowImagesParams);
    ShowImagesWindow.focus();

}

function ShowHsImages() 
{

    var LogId = document.getElementById("hLogId").value;

    var ShowImagesParams, ShowImagesURL, ShowImagesWindow, W, H, L, T;
    ShowImagesURL = 'HsShowImages.asp?Id=' + (LogId);

    W = 1000;
    H = 800;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    ShowImagesParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    ShowImagesParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    ShowImagesParams += ',width= 1000, height=800, left=' + L + ',top=' + T;

    ShowImagesWindow = window.open(ShowImagesURL, '_blank', ShowImagesParams);
    ShowImagesWindow.focus();

}

function ShowSingleImage(Data) 
{
    var ShowSingleParams, ShowSingleURL, ShowSingleWindow, W, H, L, T;
    ShowSingleURL = 'HsSingleImage.asp?Path=' + (Data);

    W = 700;
    H = 600;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    ShowSingleParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    ShowSingleParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    ShowSingleParams += ',width= 1000, height=800, left=' + L + ',top=' + T;

    ShowSingleWindow = window.open(ShowSingleURL, '_blank', ShowSingleParams);
    ShowSingleWindow.focus();
}

function ShowImageLoadChk() 
{
    // Window will close after 5 minutes
    // Works on Faults show Image 
    setTimeout("window.close()", 300000);
}



