//
// functions for MediaCo Reporting Logon & Admin
//
//################## Common

function Bypass() 
{
    window.location.replace("Autologin.asp?id=4");
}

function LogonLoadChk() 
{

    return ErrChk();
}

function PwReminder()
{
   window.location.replace("SendEmailReminder.asp?Id=" + document.getElementById("txtUserID").value);
}

function PwLoadChk(Form) 
{
  
    document.getElementById("txtUserPW").value = '';
    document.getElementById("txtUserChar").value = '';
    document.getElementById("txtPassword").value = '';

    if (Form == 'LoginPw') 
    {
        document.getElementById("txtPassword").focus();
    }

}

function PwClick(Form) 
{
    document.getElementById("txtPassword").value = '';

    if (Form == 'LoginPw') 
    {
        //document.getElementById("txtUserPW").value = '';
        //document.getElementById("txtUserChar").value = '';
        //document.getElementById("txtPassword").style.color = '#000000';
    }
}

function ResetUser(Uid) 
{

    
    var strUser = document.getElementById("LocalUserName").value;
    var ConfirmReset = window.confirm("Reset password for " + strUser);
    if (ConfirmReset == true)
    { window.location.replace("ResetUser.asp?Uid=" + Uid); }
    else
    { window.location.replace("login.asp") }    

}



function ResetPwUser() 
{
    document.getElementById("txtUserPW").value = '';
    document.getElementById("txtUserChar").value = '';
    document.getElementById("txtPassword").value = '';
    document.getElementById("btnSubmit").disabled = true;
    document.getElementById("txtPassword").focus();
}

function ResetAddPw() 
{
    document.getElementById("txtPassword").value = '';
    document.getElementById("txtUserID").value = '';
    document.getElementById("cboUsers").selectedIndex = 0;
    document.getElementById("btnSubmit").disabled = true;
}

function SetPassword(userid) 
{
    var PassWinParams, PassURL, PassWindow, W, H, L, T;
    PassURL = 'AddPw.asp?id=' + userid;

    //document.getElementById("data").style.visibility = "hidden";
    //document.getElementById("buttons").style.visibility = "hidden";

    W = 550;
    H = 400;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    PassWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    PassWinParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    PassWinParams += ',width= 550,height=400,left=' + L + ',top=' + T;

    PassWindow = window.open(PassURL, '_blank', PassWinParams);
    PassWindow.focus();
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


function DisableEnterKey(e,objId) 
{
    var key;

    if (window.event)
        key = window.event.keyCode;     //IE
    else
        key = e.which;     //firefox
        
    

    if (key == 13) 
    {
        if (objId == 'LoginPw' && document.getElementById("txtPassword").value != '')
        { return true; }
        else
        { return false; }
    }
    else 
    {
        if (objId == 'NewName')
        { return true; }
        else 
        {
            if (document.getElementById("frmName").value == 'frmLoginPw') 
            {
                if (PatternTest(key) == true) 
                {
                    var str = String.fromCharCode(key);
                    //document.getElementById("txtUserChar").value += str.replace(str, "*");
                    //document.getElementById("txtUserChar").value += String.fromCharCode(key);
                    //document.getElementById("txtUserPW").value += String.fromCharCode(key);
                    return true;                    
                }
                else 
                {
                    
                    return false;
                }
            }
            else 
            {
                if (document.getElementById("frmName").value == 'frmAddPW' || document.getElementById("frmName").value == 'frmAddUser' || document.getElementById("frmName").value == 'frmEditUser') 
                {
                    if (PatternTest(key) == true)
                    { return true; }
                    else 
                    {
                        if (key == 8)
                        { return true; }
                        else
                        { return false; }
                    }
                }
                else
                { return true; }
            }
        }            
    }
}

function PatternTest(Code)
{
    //^[a-zA-Z0-9_+-]*$ the '-' must be the last thing in case you need it too
    //var patt = /^[a-z0-9_£$#@*=+-]*$/i;
    
    var PatternToTest = /^[a-z0-9_£$#@*=+-]*$/i;
    return PatternToTest.test(String.fromCharCode(Code));
}

function EnableSubmit(EnForm) 
{
    if (document.getElementById("btnSubmit").disabled == true) 
    {
        document.getElementById("btnSubmit").disabled = false;
    }

    if (EnForm == 'LoginPw') 
    {
        var key1;
        if (window.event)
        { key1 = window.event.keyCode; }    //IE
        else
        { key1 = e.which; }

        if (key1 == 8)
        { document.getElementById("btnSubmit").disabled = true }
    }
}

function EnableSubmitAddPw() 
{
    if (document.getElementById("btnSubmit").disabled == true) 
    {
        if (document.getElementById("txtUserID").value != '')
        { document.getElementById("btnSubmit").disabled = false; }
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

//############################# Login & New User Pasword


function ValidatePW() 
{
    if (document.getElementById("txtPassword").value != '')
        { return true; }
    else 
    {
        document.getElementById("btnSubmit").disabled = true;
        alert('Password can not be blank');
        return false;
    }     
}

function SelectName(FormName) 
{
    var strId = document.getElementById("cboUsers").value;
    //document.getElementById("txtPassword").value = '';

    if (document.getElementById("cboUsers").value != '') 
    {
        document.getElementById("txtUserID").value = strId;
        //document.getElementById("txtPassword").focus();
        if (FormName != 'AddNew') 
        {            
            EnableSubmit();
        }
    }
    else 
    {
        if (FormName == 'AddNew') 
        {
            ResetAddPw();
        }      
    }
}

// Only used by login page, when user clicks name or presses enter
function SubmitLogin() 
{
    document.getElementById("frmSelectUser").submit();
}

// Only used by add password page
function TabToPw() 
{
    if (document.getElementById("cboUsers").value != '') 
    {
        document.getElementById("btnSubmit").disabled = true;
        document.getElementById("txtPassword").focus();
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

//############################ Admin Functions ###################################


//############################ Add Edit User

function UserLoad() 
{
    document.getElementById("txtUserID").focus();
}

function EditUser(Id) 
{
    window.location.replace('EditUser.asp?Uid=' + (Id));
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

    if (document.getElementById("ChkHsEdit").checked == true && document.getElementById("txtEmail").value == '') 
    {
        alert("User must have email");
        document.getElementById("txtEmail").focus();
        return false;    
    }

    if (document.getElementById("ChkHsView").checked == true && document.getElementById("txtEmail").value == '') 
    {
        alert("User must have email");
        document.getElementById("txtEmail").focus();
        return false;
    }

    if (document.getElementById("ChkNcEdit").checked == true && document.getElementById("txtEmail").value == '') 
    {
        alert("User must have email");
        document.getElementById("txtEmail").focus();
        return false;
    }

    if (document.getElementById("ChkMfEdit").checked == true && document.getElementById("txtEmail").value == '') 
    {
        alert("User must have email");
        document.getElementById("txtEmail").focus();
        return false;
    }

    if (document.getElementById("chkAdmin").checked == true && document.getElementById("txtEmail").value == '') 
    {
        alert("User must have email");
        document.getElementById("txtEmail").focus();
        return false;
    }
    

    //    if (document.getElementById("chkActive").checked == true && document.getElementById("frmName").value == 'frmAddUser')
    if (document.getElementById("chkActive").checked == true) 
    {
        var nobox = 0;
        if (document.getElementById("chkHS").checked == false)
        { nobox += 1; }

        if (document.getElementById("chkNC").checked == false)
        { nobox += 1; }

        if (document.getElementById("chkMF").checked == false)
        { nobox += 1; }
        
        if (nobox == 3) 
        {
            alert("Select at least one reporting option");
            return false;
        }
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

//############################ Add Edit Email Settings


function EditEmailUser(Id) 
{
    window.location.replace('EditEmailUser.asp?Uid=' + (Id));
}

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
        alert("Select a user");
        document.getElementById("cboEmailUser").focus();
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

function PickEmailUser() 
{

    var strPicked = document.getElementById("cboEmailUser").value;
    var e = document.getElementById("cboEmailUser");
    var strTxt = e.options[e.selectedIndex].text;

    if (document.getElementById("cboEmailUser").value != '') 
    {
        document.getElementById("hEmailUserId").value = strPicked;
        document.getElementById("txtUserID").value = strTxt;
        document.getElementById("btnSubmit").disabled == false;
    }
    else 
    {
        document.getElementById("hEmailUserId").value = '0'
        document.getElementById("txtUserID").value = "";
        document.getElementById("btnSubmit").disabled == true;
    }



}








