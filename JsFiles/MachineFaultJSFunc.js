//
// functions for Machine Faults
//
//################## Common

function ErrChk() {
    var emsg = document.getElementById("ErrBox").value
    if (document.getElementById("ErrBox").value.length > 1)
    { window.alert(emsg); }
}

function LogOff() 
{
    window.location.replace("LogOff.asp");
}

function GoBackOption() 
{
    window.location.replace('MfSelectOption.asp');
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
}
   

function ResetPage() 
{
    window.location.reload(true);
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
        else 
        {
            window.location.replace("MfEditLog.asp?id=" + ordertoget)
        }
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

function RelocateMachine()
{
    var strId = document.getElementById("cboMachine").value;
    if (document.getElementById("cboMachine").value != '') 
    {
        document.getElementById("Machine").value = strId;
        window.location.replace("MfDisplayMachine.asp?id=" + strId);
    }
    else {
        document.getElementById("Machine").value = '';
    }
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

//############################ Add Edit Log

function FaultSelectAdd() 
{
    var strFiles

    if (document.getElementById("cboFaultGroup").value != '') 
    {
        strFiles = document.getElementById("cboFaultGroup").value;
        var e = document.getElementById("cboFaultGroup");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("hSelectId").value;
        var TempText = document.getElementById("Grouplist").innerHTML;

        var pos = StrToSearch.search(strFiles);        
        if (pos == -1) 
        {
            if (document.getElementById("hSelectId").value == '') 
            {
                document.getElementById("hGroupId").value = Number(strFiles);
                document.getElementById("hSelectId").value = strFiles;
                document.getElementById("Grouplist").innerHTML = strTxt;
                document.getElementById("GroupTitle").innerHTML = 'Selected Fault Groups';
            }
            else 
            {
                document.getElementById("hGroupId").value += ',' + Number(strFiles);
                document.getElementById("hSelectId").value += ',' + strFiles;
                document.getElementById("Grouplist").innerHTML = TempText + '<br />' + strTxt;
                document.getElementById("GroupTitle").innerHTML = 'Selected Fault Groups';
            }
            //document.getElementById("Grouplist").style.color = "#009933";
        }
    }
    else 
    {
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
        document.getElementById("Grouplist").innerHTML = '';
        document.getElementById("GroupTitle").innerHTML = '';
        //document.getElementById("Grouplist").style.color = "#ffffff";
    }
}

function FaultSelectEdit() 
{
    var strFiles

    if (document.getElementById("cboFaultGroup").value != '') 
    {
        strFiles = document.getElementById("cboFaultGroup").value;
        var e = document.getElementById("cboFaultGroup");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("hSelectId").value;
        var TempText = document.getElementById("Grouplist").innerHTML;

        var pos = StrToSearch.search(strFiles);
        if (pos == -1) 
        {
            if (document.getElementById("hGroupId").value == '') 
            {
                document.getElementById("hGroupId").value = Number(strFiles);
                document.getElementById("hSelectId").value = strFiles;
                document.getElementById("Grouplist").innerHTML = strTxt;
            }
            else 
            {
                document.getElementById("hGroupId").value += ',' + Number(strFiles);
                document.getElementById("hSelectId").value += ',' + strFiles;
                document.getElementById("Grouplist").innerHTML = TempText + ', ' + strTxt;
            }
            //document.getElementById("Grouplist").style.color = "#FF0000";
        }
    }
    else 
    {
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
        document.getElementById("Grouplist").innerHTML = '';
        //document.getElementById("Grouplist").style.color = "#ffffff";
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
        document.getElementById("txtError").focus();
    }
}

function MachineMouseMove() 
{
//    if (document.getElementById("hType").value != '')
//    {alert('Changing machine will reset any selected items\nSubject & notes will remain') }
}

function AddMachine() 
{    
    var MachineId = document.getElementById("cboMachine").value;
    var MachineArr = MachineId.split("#");
    
    if (document.getElementById("cboMachine").value != '')
    {
        document.getElementById("hMachine").value = MachineArr[0];
        document.getElementById("hType").value = MachineArr[1];
        document.getElementById("hGroups").value = MachineArr[2];
        document.getElementById("chkRecurring").disabled = false;
        document.getElementById("chkFixed").disabled = false;
        document.getElementById("cboSeverity").selectedIndex = 0;
        document.getElementById("cboSeverity").disabled = false;
        document.getElementById("cboFaultGroup").disabled = false;
        document.getElementById("cboFaultGroup").selectedIndex = 0;
        document.getElementById("Grouplist").innerHTML = '';
        document.getElementById("hGroupId").value = '';
        document.getElementById("GroupTitle").innerHTML = '';
        document.getElementById("hSeverity").value = '';
        document.getElementById("chkRecurring").checked = false;
        document.getElementById("chkFixed").checked = false;
        document.getElementById("hSelectId").value = '';
        
        getFaultGroup(document.getElementById("hGroups").value);       
    }
    else 
    {
        document.getElementById("hMachine").value = '';
        document.getElementById("hType").value = '';
        document.getElementById("hGroups").value = '';
        document.getElementById("chkRecurring").checked = false;
        document.getElementById("chkRecurring").disabled = true;
        document.getElementById("chkFixed").checked = false;
        document.getElementById("chkFixed").disabled = true;
        document.getElementById("cboSeverity").selectedIndex = 0;
        document.getElementById("cboSeverity").disabled = true;
        document.getElementById("cboFaultGroup").innerHTML = "<option value=''>Select Fault Group</option>";
        document.getElementById("cboFaultGroup").selectedIndex = 0;
        document.getElementById("cboFaultGroup").disabled = true;
        document.getElementById("hMachine").value = '';
        document.getElementById("hSeverity").value = '';
        document.getElementById("hType").value = '';
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
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

function EditLogLoad() 
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

        getFaultGroup(document.getElementById("hGroups").value);
    }
}


function MFLoadEditOrder(EditId) 
{

    var Sender = document.getElementById("Sendpage").value;
    window.location.replace("MfEditLog.asp?id=" + EditId + "&sp=" + Sender);

}


function MFValidateLog(SendingPage)
{
    var ValidateMsg = ''
    var ValidateError = 0

    if (SendingPage == 'Add') 
    {
//        if (document.getElementById("txtDesc").value == '') 
//        {
//            ValidateMsg = "Subject is missing\n";
//            ValidateError += 1
//        }

        if (document.getElementById("txtError").value == '') 
        {
            ValidateMsg += "Enter some details\n";
            ValidateError += 1
        }

        if (document.getElementById("hMachine").value == '') 
        {
            ValidateMsg += "Select a machine\n";
            ValidateError += 1
        }

        if (document.getElementById("hSeverity").value == '') 
        {
            ValidateMsg += "Select severity of fault\n";
            ValidateError += 1
        }

        if (document.getElementById("hGroupId").value == '') 
        {
            ValidateMsg += "Select a fault group\n";
            ValidateError += 1
        }
    }
    else 
    {
        if (document.getElementById("txtRepair").innerHTML = '') 
        {
            ValidateMsg = "Enter some repair notes\n";
            ValidateError += 1
        }

        if (document.getElementById("hSeverity").value == '') 
        {
            ValidateMsg += "Select severity of fault\n";
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
    {UploadURL = "MfUploadEdit.asp"; }
    else
    {UploadURL = "MfUploadAdd.asp"; }
        
    UploadWindow = window.open(UploadURL, '_blank', UploadWinParams);

}

function LoadImage(Id) 
{


    var ShowImagesParams, ShowImagesURL, ShowImagesWindow, W, H, L, T;
    ShowImagesURL = 'MfShowImages.asp?Id=' + (Id);

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

function ShowImages() 
{

    var LogId = document.getElementById("hLogId").value;
      
    var ShowImagesParams, ShowImagesURL, ShowImagesWindow, W, H, L, T;
    ShowImagesURL = 'MfShowImages.asp?Id=' + (LogId);   

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
    ShowSingleURL = 'MfSingleImage.asp?Path=' + (Data);

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
 
 


//############################ Admin Functions ###################################

function MachinePageLoad(FormName)
{

    if (FormName == 'Edit') 
    {
        getFaultGroup('#' + document.getElementById("hMachineType").value);
    }
    else 
    {
        document.getElementById("txtMachineID").focus();
        document.getElementById("cboMachineType").disabled = true;
    }       
}

function AddMachineEnableType()
{
    document.getElementById("cboMachineType").disabled = false;
}

function SelectMachine() 
{
    var strId = document.getElementById("cboMachine").value;
    if (document.getElementById("cboMachine").value != '') 
    {
        document.getElementById("hMachineId").value = strId;
        document.getElementById("btnsubmit").disabled = false;
    }
    else 
    {
        document.getElementById("hMachineId").value = '';
        document.getElementById("btnsubmit").disabled = true;
    }
}

function MachineNameChanged() 
{
    var NameChange = Number(document.getElementById("hMachineChange").value)
    NameChange += 1;
    document.getElementById("hMachineChange").value = NameChange;   

}

function MachineType() 
{
    var strMachineTypeId = document.getElementById("cboMachineType").value;
    if (document.getElementById("cboMachineType").value != '') 
    {
        document.getElementById("hMachineType").value = strMachineTypeId;
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
        document.getElementById("Grouplist").innerHTML = '';
        document.getElementById("cboFaultGroup").disabled = false;
        document.getElementById("cboFaultGroup").selectedIndex = 0;
        getFaultGroup('#' + document.getElementById("hMachineType").value);              
    }
    else 
    {
        document.getElementById("hMachineType").value = '';
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
        document.getElementById("Grouplist").innerHTML = '';
        document.getElementById("cboFaultGroup").selectedIndex = 0;
        document.getElementById("cboFaultGroup").disabled = true;
    }

}

function FaultSelectNewMachine() 
{
    var strNewMachine

    if (document.getElementById("cboFaultGroup").value != '') 
    {
        strNewMachine = document.getElementById("cboFaultGroup").value;
        var e = document.getElementById("cboFaultGroup");
        var strTxt = e.options[e.selectedIndex].text;
        var StrToSearch = document.getElementById("hSelectId").value;
        var TempText = document.getElementById("Grouplist").innerHTML;

        var pos = StrToSearch.search(strNewMachine);        
        if (pos == -1) 
        {
            if (document.getElementById("hSelectId").value == '') 
            {
                document.getElementById("hGroupId").value = Number(strNewMachine);
                document.getElementById("hSelectId").value = strNewMachine;
                document.getElementById("Grouplist").innerHTML = strTxt;
            }
            else 
            {
                document.getElementById("hGroupId").value += ',' + Number(strNewMachine);
                document.getElementById("hSelectId").value += ',' + strNewMachine;
                document.getElementById("Grouplist").innerHTML = TempText + ', ' + strTxt;
            }
        }
    }
    else 
    {
        document.getElementById("hGroupId").value = '';
        document.getElementById("hSelectId").value = '';
        document.getElementById("Grouplist").innerHTML = '';
    }
}


function ValidateMachine(FormName) 
{

    var MachineCount = 0
    var MachineTxt = document.getElementById("txtMachineID").value.toUpperCase()
    var MachineChk = document.getElementById("hMachineList").value.toUpperCase()
    var MachineArr = MachineChk.split("#");

    if (document.getElementById("hMachineChange").value == '0' && FormName == "Edit")
    { }  // do Nothing
    else
        for (MachineCount = 0; MachineCount < MachineArr.length; MachineCount++) 
        {
            if (MachineTxt == MachineArr[MachineCount]) 
            {
            alert("Machine Name in use, enter a new name");
            return false;
        }
    }

    if (document.getElementById("txtMachineID").value == '') 
    {
        alert("Enter a name");
        document.getElementById("txtMachineID").focus();
        return false
    }

    if (document.getElementById("hMachineType").value == '') 
    {
        alert("Select a Machine Type");
        return false
    }

    if (document.getElementById("hGroupId").value == '' || document.getElementById("hSelectId").value == '') 
    {
        if (document.getElementById("hMachineChange").value == '0' && FormName == "Edit")
        { return true}  // do nothing deactivating machine
        else 
        {
            alert("Select some groups");
            return false
        }
    }
    
}

//########################## Group

function MachineTypeGroup() 
{
    var strMachineTypeId = document.getElementById("cboMachineType").value;
    if (document.getElementById("cboMachineType").value != '') 
    {
        document.getElementById("hMachineType").value = strMachineTypeId;
        document.getElementById("hGroupId").value = '';
        document.getElementById("cboFaultGroup").disabled = false;
        getFaultGroup(document.getElementById("hMachineType").value, "False");        
    }
    else 
    {
        document.getElementById("hMachineType").value = '';
        document.getElementById("cboFaultGroup").disabled = true;
        document.getElementById("hGroupId").value = '';
    }

}

function FaultSelect()
{
    var strGroupId = document.getElementById("cboFaultGroup").value;
    if (document.getElementById("cboFaultGroup").value != '') 
    {
        document.getElementById("hGroupId").value = strGroupId;
        document.getElementById("btnsubmit").disabled = false;
    }
    else 
    {
        document.getElementById("hGroupId").value = '';
        document.getElementById("btnsubmit").disabled = true;
    }
}

function GroupNameChanged() 
{
    var NameChange = Number(document.getElementById("hGroupChange").value)
    NameChange += 1;
    document.getElementById("hGroupChange").value = NameChange;

}

function GroupType() 
{
    var strGroupTypeId = document.getElementById("cboGroupType").value;
    if (document.getElementById("cboGroupType").value != '') 
    {
        document.getElementById("hGroupType").value = strGroupTypeId;
    }
    else 
    {
        document.getElementById("hGroupType").value = '';
    }

}

function ValidateGroup(FormName) 
{

    var GroupCount = 0
    var GroupTxt = document.getElementById("txtGroupID").value.toUpperCase()
    var GroupChk = document.getElementById("hGroupList").value.toUpperCase()
    var GroupArr = GroupChk.split("#");

    if (document.getElementById("hGroupChange").value == '0' && FormName == "Edit")
    { }  // do Nothing
    else
        for (GroupCount = 0; GroupCount < GroupArr.length; GroupCount++) 
        {
            if (GroupTxt == GroupArr[GroupCount]) {
                alert("Group Name in use, enter a new name");
            return false;
        }
    }

    if (document.getElementById("txtGroupID").value == '') 
    {
        alert("Enter a name");
        document.getElementById("txtGroupID").focus();
        return false
    }

    if (document.getElementById("hGroupType").value == '') 
    {
        alert("Select a Group Type");
        return false
    }

}

