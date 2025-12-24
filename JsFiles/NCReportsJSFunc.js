//
// functions for NC Reporting
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

function GoBackNCOption() 
{
    window.location.replace('NcSelectOption.asp');
}

function ShowOrder(e) 
{
//    var upkey;

//    if (window.event)
//    { upkey = window.event.keyCode; }     //IE
//    else
//    { upkey = e.which; }    //firefox

//    if (upkey == 13) 
//    {
//        var ordertoget = document.getElementById("JumpTxt").value;
//        if (isNaN(ordertoget) || ordertoget == '' || ordertoget == '0')
//        //if (rowtoget != '')
//        {
//            alert('Enter a valid number')
//        }
//        else {
//            window.location.replace("EditLog.asp?id=" + ordertoget + "&sp=0");
//        }
//    }
}

function DisableEnterKey(e, objId) {
    var key;

    if (window.event)
        key = window.event.keyCode;     //IE
    else
        key = e.which;     //firefox

    if (key == 13) {
        return false;
    }

}

function PatternTest(Code)
{
    //^[a-zA-Z0-9_+-]*$ the '-' must be the last thing in case you need it too
    //var patt = /^[a-z0-9_£$#@*=+-]*$/i;
    
    var PatternToTest = /^[a-z0-9_£$#@*=+-]*$/i;
    return PatternToTest.test(String.fromCharCode(Code));
}

function EnableSubmit() 
{
    if (document.getElementById("btnSubmit").disabled == true) 
    {
        document.getElementById("btnSubmit").disabled = false;
    }
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

function OptionsLoad() 
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

//############################ Add Edit

function AddItem() 
{
    var strItem

    if (document.getElementById("cboItem").value != '') 
    {
        strItem = document.getElementById("cboItem").value;
        var e = document.getElementById("cboItem");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("hItemsAlpha").value
        var TempText = document.getElementById("SelectedItem").value

        document.getElementById("cboDeptSelect").disabled = false;

        var pos = StrToSearch.search(strTxt);
        if (pos == -1) 
        {
            if (document.getElementById("hItemsAlpha").value == '') 
            {
                document.getElementById("hItemsAlpha").value = strTxt;
                document.getElementById("SelectedItem").value = strTxt;
                document.getElementById("hItemId").value = strItem;
            }
            else {
                document.getElementById("hItemsAlpha").value += ',' + strTxt;
                document.getElementById("SelectedItem").value = TempText + ', ' + strTxt;
                document.getElementById("hItemId").value += ',' + strItem;
                
            }
        }
    }
    else 
    {
        document.getElementById("hItemsAlpha").value = '';
        document.getElementById("SelectedItem").value = ''
        document.getElementById("hItemId").value = '';
        document.getElementById("cboReason").disabled = true;
        
        
//        document.getElementById("cboDeptSelect").selectedIndex = 0;
//        document.getElementById("cboDeptSelect").disabled = true;
//        document.getElementById("cboReason").selectedIndex = 0;
//        document.getElementById("cboReason").disabled = true;

//        document.getElementById("hDept").value = '';
//        document.getElementById("SelectedDept").value = '';
//        document.getElementById("hGroup").value = '';
//        document.getElementById("hDeptSelected").value = '';
//        document.getElementById("hGroupSelected").value = '';

//        document.getElementById("hReason").value = '';
//        document.getElementById("hReasonSelected").value = '';
//        document.getElementById("SelectedReason").value = '';
//        document.getElementById("txtDetails").disabled = true;
    }
}


function AddDept() 
{
    var strDept;
    

    if (document.getElementById("cboDeptSelect").value != '') 
    {
        strDept = document.getElementById("cboDeptSelect").value;

        var DeptChk = strDept;
        var DeptArray = DeptChk.split("#");
        
        var e = document.getElementById("cboDeptSelect");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("SelectedDept").value;
        var TempText = document.getElementById("SelectedDept").value;

        document.getElementById("cboReason").disabled = false;

        // Work Centres
        var pos = StrToSearch.search(strTxt);
        if (pos == -1) 
        {
            if (document.getElementById("hDept").value == '') 
            {
                document.getElementById("hDeptSelected").value = DeptArray[1];    
                document.getElementById("hDept").value = Number(DeptArray[1]);
                document.getElementById("SelectedDept").value = strTxt;
            }
            else 
            {
                document.getElementById("hDeptSelected").value += ',' + DeptArray[1]; 
                document.getElementById("hDept").value += ',' + Number(DeptArray[1]);
                document.getElementById("SelectedDept").value = TempText + ', ' + strTxt;
            }

            if (document.getElementById("hGroupSelected").value == '') {
                document.getElementById("hGroup").value = Number(DeptArray[0]);
                document.getElementById("hGroupSelected").value = DeptArray[0];
            }
            else {
                document.getElementById("hGroup").value += ',' + Number(DeptArray[0]);
                document.getElementById("hGroupSelected").value += ',' + DeptArray[0];
            }
            getReasons(document.getElementById("hGroup").value);
            
        }
        
    }
    else 
    {
        document.getElementById("hDept").value = '';
        document.getElementById("SelectedDept").value = '';
        document.getElementById("hGroup").value = '';
        document.getElementById("hDeptSelected").value = '';
        document.getElementById("hGroupSelected").value = '';
        document.getElementById("cboReason").disabled = true;
    }
}

function AddGroupNoJob() 
{
    var strGroup;

    // Only uses group

    if (document.getElementById("cboGroupSelect").value != '') 
    {
        strGroup = document.getElementById("cboGroupSelect").value;
        
        var e = document.getElementById("cboGroupSelect");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("SelectedGroup").value;
        var TempText = document.getElementById("SelectedGroup").value;

        document.getElementById("cboReason").disabled = false;

        var pos = StrToSearch.search(strTxt);
        if (pos == -1) 
        {
            if (document.getElementById("hGroup").value == '') 
            {
                document.getElementById("hGroupSelected").value = strGroup;
                document.getElementById("hGroup").value = Number(strGroup);  
                document.getElementById("SelectedGroup").value = strTxt;
            }
            else 
            {
                document.getElementById("hGroupSelected").value += ',' + strGroup
                document.getElementById("hGroup").value += ',' + Number(strGroup);  
                document.getElementById("SelectedGroup").value = TempText + ', ' + strTxt;
            }
            getReasons(document.getElementById("hGroup").value);
        }
    }
    else 
    {
        document.getElementById("hGroup").value = '';
        document.getElementById("SelectedGroup").value = '';
        document.getElementById("hGroupSelected").value = '';
        document.getElementById("cboReason").disabled = true;
    }
}


function AddReason() 
{
    var strReason;

    if (document.getElementById("cboReason").value != '') 
    {
        //document.getElementById("hReasonId").value = ReasonId
        //document.getElementById("hIssue").value = strProblem;


        strReason = document.getElementById("cboReason").value;

        var e = document.getElementById("cboReason");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("SelectedReason").value;
        var TempText = document.getElementById("SelectedReason").value;

        document.getElementById("txtDetails").disabled = false;

        var pos = StrToSearch.search(strTxt);
        if (pos == -1) 
        {
            if (document.getElementById("hReason").value == '') 
            {
                document.getElementById("hReason").value = Number(strReason);
                document.getElementById("hReasonSelected").value = strReason;
                document.getElementById("SelectedReason").value = strTxt;
            }
            else 
            {

                document.getElementById("hReason").value += ',' + Number(strReason);
                document.getElementById("hReasonSelected").value += ',' + strReason; 
                document.getElementById("SelectedReason").value = TempText + ', ' + strTxt;
            }
        }
    }
    else 
    {
        document.getElementById("hReason").value = '';
        document.getElementById("hReasonSelected").value = '';
        document.getElementById("SelectedReason").value = '';
        document.getElementById("txtDetails").disabled = true;
    }
}

// ################################## Edit ###################################

// Will need EditItem, EditDept & EditReason
function FaultSelectEdit() 
{
    var strFiles

    if (document.getElementById("cboFaultGroup").value != '') 
    {
        strFiles = document.getElementById("cboFaultGroup").value;
        var e = document.getElementById("cboFaultGroup");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("hGroupId").value;
        var TempText = document.getElementById("Grouplist").innerHTML;

        var pos = StrToSearch.search(strFiles);
        if (pos != 0) 
        {
            if (document.getElementById("hGroupId").value == '') 
            {
                document.getElementById("hGroupId").value = strFiles
                document.getElementById("Grouplist").innerHTML = strTxt;
            }
            else 
            {
                document.getElementById("hGroupId").value += ',' + strFiles;
                document.getElementById("Grouplist").innerHTML = TempText + ', ' + strTxt;
            }
        }
    }
    else {
        document.getElementById("hGroupId").value = '';
        document.getElementById("Groupelist").innerHTML = '';
    }
}


function AddLogLoad()
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





function Severity() 
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

function AddLocation() 
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

function NCEditLogLoad() 
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




function NcEditRecord(EditId, EditType)
 {

    
    var EditLocation = ""
    if (EditType == Number(0))
    { EditLocation = "NcEditRecordNJ.asp?id=" + EditId; }
    else
    { EditLocation = "NcEditRecordCL.asp?id=" + EditId; }



    window.location.replace(EditLocation);
    //alert(EditLocation);

}

function NcViewRecordUser(EditId, EditType)
{
    var ViewLocation = ""
    if (EditType == Number(0))
    { ViewLocation = "NcViewUserLogNJ.asp?id=" + EditId; }
    else
    { ViewLocation = "NcViewUserLogCL.asp?id=" + EditId; }
    
    window.location.replace(ViewLocation);

}


function NCValidateLogNJ(SendingPage)
{
    var ValidateMsg = ''
    var ValidateError = 0

    if (SendingPage == 'Add') 
    {
        

        if (document.getElementById("hGroup").value == '') {
            ValidateMsg += "Select a department\n";
            ValidateError += 1
        }

        if (document.getElementById("hReason").value == '') 
        {
            ValidateMsg += "Select a reason\n";
            ValidateError += 1
        }

        if (document.getElementById("txtDetails").value == '') {
            ValidateMsg += "Enter some details\n";
            ValidateError += 1
        }
       
    }
    else 
    {
        
        if (document.getElementById("txtResponse").value == '') 
        {
            ValidateMsg += "Enter some comments\n";
            ValidateError += 1
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

function NCValidateLogCL(SendingPage) 
{
    var ValidateMsg = ''
    var ValidateError = 0

    if (SendingPage == 'Add') 
    {
        

        if (document.getElementById("hItemId").value == '') {
            ValidateMsg += "Select a Item\n";
            ValidateError += 1
        }

        if (document.getElementById("hDept").value == '') {
            ValidateMsg += "Select a work centre\n";
            ValidateError += 1
        }
        
        if (document.getElementById("hReason").value == '') {
            ValidateMsg += "Select a reason\n";
            ValidateError += 1
        }

        if (document.getElementById("txtDetails").value == '') {
            ValidateMsg += "Enter some details\n";
            ValidateError += 1
        }
        

    }
    else 
    {
        if (document.getElementById("hHasNotes").value == 'False') {
            if (document.getElementById("txtResponse").value == '') {
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
    else {
        return true;
    }
}
//############################ Admin Functions ###################################



//########################## Group


//function FaultSelect()
//{
//    var strGroupId = document.getElementById("cboFaultGroup").value;
//    if (document.getElementById("cboFaultGroup").value != '') 
//    {
//        document.getElementById("hGroupId").value = strGroupId;
//        document.getElementById("btnsubmit").disabled = false;
//    }
//    else 
//    {
//        document.getElementById("hGroupId").value = '';
//        document.getElementById("btnsubmit").disabled = true;
//    }
//}

function ReasonNameChanged() 
{
    var NameChange = Number(document.getElementById("hReasonChange").value)
    NameChange += 1;
    document.getElementById("hReasonChange").value = NameChange;

}

function ReasonSelect() 
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

function ReasonTypeSelect() 
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

function ValidateReason(FormName) 
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

//############################ Add Edit User

function UserLoad() 
{
    document.getElementById("txtUserID").focus();
}

function EditUser(Id) 
{
    window.location.replace('EditUser.asp?Uid=' + (Id));
}

function EditEmailUser(Id) 
{
    window.location.replace('EditEmailUser.asp?Uid=' + (Id));
}


function UserNameChanged() 
{
    var NameChange = Number(document.getElementById("hNameChange").value)
    NameChange += 1;
    document.getElementById("hNameChange").value = NameChange;

}

function DelUser() 
{

    if (document.getElementById("hDelete").value == "0") 
    {
        document.getElementById("hDelete").value = "1";
        document.getElementById("btnDel").value = "Restore";
        //document.getElementById("Deltxt").innerHTML = "Restore User";
    }
    else    
    {
        document.getElementById("hDelete").value = "0";
        document.getElementById("btnDel").value = "Delete";
        //document.getElementById("Deltxt").innerHTML = "Delete User";
    }
}

function ValidateUser(FormName) 
{

    var UserCount 
    var UserTxt 
    var UserChk
    var UserArr

    if (FormName == "Edit") 
    {
        if (document.getElementById("hDelete").value == "1")
        { return true }
    }  
    else
    {    
        UserCount = 0
        UserTxt = document.getElementById("txtUserID").value.toUpperCase()
        UserChk = document.getElementById("hUserList").value.toUpperCase()
        UserArr = UserChk.split("#");
    
        for (UserCount = 0; UserCount < UserArr.length; UserCount++) 
        {
            if (UserTxt == UserArr[UserCount]) 
            {
                alert("User Name in use, enter a new name");
                return false;
            }
        }
    }

    if (document.getElementById("txtUserID").value == '') 
    {
        alert("Enter a user name");
        document.getElementById("txtUserID").focus();
        return false;
    }

    if (document.getElementById("txtPassword").value == '') 
    {
        if (document.getElementById("chkActive").checked == true) 
        {
            alert("Enter a password");
            document.getElementById("txtPassword").focus();
            return false;
        }
        else
        { return true }        
    }
    else 
    {
         return true;
    }
}


//############################ Add Edit Email User


function ValidateEmailUser(FormName) 
{

    var UserCount = 0
    var UserTxt = document.getElementById("txtUserID").value.toUpperCase()
    var UserChk = document.getElementById("hUserList").value.toUpperCase()
    var UserArr = UserChk.split("#");

    if (document.getElementById("hNameChange").value == '0' && FormName == "Edit")
    { }  // do Nothing
    else
        for (UserCount = 0; UserCount < UserArr.length; UserCount++) 
        {
            if (UserTxt == UserArr[UserCount]) 
        {
            alert("User Name in use, enter a new name");
            return false;
        }
    }

    if (document.getElementById("txtUserID").value == '') 
    {
        alert("Enter a user name");
        document.getElementById("txtUserID").focus();
        return false
    }

    if (document.getElementById("txtEmail").value == '') 
    {
        alert("Enter a email address");
        document.getElementById("txtEmail").focus();
        return false
    }
    else
    {
        if (isValidEmail(document.getElementById('txtEmail').value) == false) 
        {
            alert("Please enter a valid Email Address");
            document.getElementById("txtEmail").focus();
            return false
        }
        else
        { return true; }
    }
}

function isValidEmail(emailAddress) 
{
    // change 2,4 to 2, 5 6 7 8 or what ever to allow long names like museum. currently set at 4
    var emailRegExp = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,4})+$/
    return emailRegExp.test(emailAddress);
}



