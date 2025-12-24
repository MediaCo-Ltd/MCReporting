// Functions for Clarity Reports 

var mins;
var secs;



//window.onload = init;
//done in body onload ="javascript:init();"
//style="visibility: hidden"

function CountDown()
{
    mins = 1 * m("15"); // change minutes here, default 15
    secs = 0 + s(":01"); // change seconds here (always add an additional second to your total)
    redo();
}

function m(obj)
{
    for (var i = 0; i < obj.length; i++) 
    {
        if (obj.substring(i, i + 1) == ":")
            break;
    }
    return (obj.substring(0, i));
}

function s(obj)
{
    for (var i = 0; i < obj.length; i++)
    {
        if (obj.substring(i, i + 1) == ":")
            break;
    }
    return (obj.substring(i + 1, obj.length));
}

function dis(mins, secs)
{
    var disp;
    
    if (mins <= 9) 
    {
        disp = " 0"; // put to " 0" if want 09        
    }
    else
    {
        disp = " ";        
    }
    disp += mins + ":";
    if (secs <= 9)
    {
        disp += "0" + secs;        
    }
    else
    {
        disp += secs;        
    }
    return (disp);
}

function redo()
{
    secs--;
    if (secs == -1)
    {
        secs = 59;
        mins--;
    }
    
    document.getElementById("CountDown").innerHTML = 'Update&nbsp;in&nbsp;' + dis(mins, secs);

    var thisDate = new Date();
    var hours = thisDate.getHours();
    var minutes = thisDate.getMinutes();
    var seconds = thisDate.getSeconds();

    if (document.getElementById("frmName").value == 'RequiredToday') 
    {
        if ((hours == 15) && (minutes == 30) && (seconds == 30)) 
        {
            if (mins > 1) 
            {
            setTimeout("window.location.reload(true)", 1000);
            }
        }
    }


    if ((mins == 0) && (secs == 0)) 
    {
        document.getElementById("CountDown").innerHTML = '';  //'Updating';
        document.getElementById("ProgBar1").width = '111';
        document.getElementById("ProgBar1").height = '11';
        document.getElementById("ProgBar1").src = 'Images/prog_9.GIF';
        document.getElementById("ProgBar1").style.visibility = "visible";
        document.body.style.cursor = 'wait';

        //setTimeout("window.location.reload()", 1000);
        setTimeout("window.location.reload(true)", 1000);
    }
    else 
    {
        CountDown = setTimeout("redo()", 1000);
    }
}

function init()
{
    
    document.getElementById("ProgBar1").style.visibility = "hidden";
    document.getElementById("ProgBar1").width = '.1';
    document.getElementById("ProgBar1").height = '.1';
    document.body.style.cursor = 'default';
    //document.getElementById("JumpTxt").focus();
    if ((document.getElementById("frmName").value == 'NoLabour') && (document.getElementById("hReqStatus").value == 'True'))
    { document.getElementById("user7").style.visibility = "visible"; }   
    CountDown();
}

function restorecolour(TrId) 
{
    document.getElementById('Td' + TrId).style.backgroundColor = document.getElementById(TrId).style.backgroundColor   
}


function hide(IcId) 
{

    if (document.getElementById(IcId).style.display == "none") 
    {
        document.getElementById(IcId).style.display = "inherit";
        document.getElementById('msg' + IcId).setAttribute("title", "Click to hide");
        document.getElementById('msg' + IcId).style.color = 'black';
    }
    else 
    {
        document.getElementById(IcId).style.display = "none";
        document.getElementById('msg' + IcId).setAttribute("title", "Click to show");
        document.getElementById('msg' + IcId).style.color = 'red';
    }
}

function closeall() 
{

    var myList = document.getElementsByClassName("Div");
    var myMsg = document.getElementsByClassName("Msg");

    if (document.getElementById("Collapse").value == "Collapse") 
    {
        for (var i = 0; i < myList.length; i++) 
        {
            myList[i].style.display = "inherit";
        }

        for (var i = 0; i < myMsg.length; i++) 
        {
            myMsg[i].style.color = 'black';
            myMsg[i].setAttribute("title", "Click to hide");
        }
        document.getElementById("Collapse").value = "";       
        

        if (document.getElementById("frmType").value == "DPCBW") 
        {
            document.getElementById("collapseMsg").style.color = '#000000';
            document.getElementById("collapseMsg").setAttribute("title", "Hides all data, just leaving the codes\nAllowing you to expand the ones you want");
        }
        else 
        {
            if (document.getElementById("frmType").value == "DPC") 
            {
                document.getElementById("collapseMsg").style.color = '#0069AA';
                document.getElementById("collapseMsg").setAttribute("title", "Hides all data, just leaving the codes\nAllowing you to expand the ones you want");
            }
            else 
            {
                document.getElementById("collapseMsg").style.color = '#0069AA';
                document.getElementById("collapseMsg").setAttribute("title", "Hides all data, just leaving the dates\nAllowing you to expand the ones you want");
            }
        }
        document.getElementById("collapseMsg").innerHTML = 'Click to hide all';
    }
    else 
    {
        for (var i = 0; i < myList.length; i++) 
        {
            myList[i].style.display = "none";
        }

        for (var i = 0; i < myMsg.length; i++) 
        {
            myMsg[i].style.color = 'red';
            myMsg[i].setAttribute("title", "Click to show");
        }
        document.getElementById("Collapse").value = "Collapse";
        document.getElementById("collapseMsg").style.color = '#FF0000';
        document.getElementById("collapseMsg").innerHTML = 'Click to show all';
        document.getElementById("collapseMsg").setAttribute("title", "Show all data");
    }
}


function TabToNext(e, objId) 
{
    var upkey;

    if (window.event)
    { upkey = window.event.keyCode; }     //IE
    else
    { upkey = e.which; }      //firefox

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

function SetEstPick(QId) 
{
    var DetailWinParams, DetailURL, DetailWindow, W, H, L, T;
    W = 400;
    H = 400;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    DetailWinParams = 'toolbar=no,location=no,status=no,menubar=no,copyhistory=no';
    DetailWinParams += ',scrollbars=no,statusbar=no,resizable=yes';
    DetailWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=' + T + '';

    DetailURL = 'SetEstPick.asp?Id=' + (QId);
    DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);
    //window.location.replace (DetailURL)

}

function RedoNumbers(QuoteId) 
{
    var DetailWinParams, DetailURL, DetailWindow, W, H, L, T;

    var RedoWebLocation;

    if (document.getElementById("Location").value == "Work") 
    {
        // Using Qc web site has session, but is abondoned in web
        // RedoWebLocation = 'http://192.168.20.1/Qc/Rl.asp?QuoteId=' + (QuoteId);
        // Using Own web site no session

        //RedoWebLocation = 'http://192.168.20.1/Redo/?QuoteId=' + (QuoteId);
        
        // If File-Server Falls over link to MC-Sql2012 or ClarityDemo
        // RedoWebLocation = 'http://192.168.20.205/Redo/?QuoteId=' + (QuoteId);
        // ClarityDemo
        RedoWebLocation = 'http://192.168.20.85/Redo/?QuoteId=' + (QuoteId);
    }

    if (document.getElementById("Location").value == "Home") 
    {
        RedoWebLocation = 'Redo.asp?QuoteId=' + (QuoteId);
    }

    if (document.getElementById("Location").value == "WorkTest") 
    {
        // Local copy of data
        RedoWebLocation = 'Redo.asp?QuoteId=' + (QuoteId);
        // File-server
        // RedoWebLocation = 'http://192.168.20.1/Redo/?QuoteId=' + (QuoteId);
        // MC-SQL2012
        // RedoWebLocation = 'http://192.168.20.205/Redo/?QuoteId=' + (QuoteId);
        // ClarityDemo
        // RedoWebLocation = 'http://192.168.20.85/Redo/?QuoteId=' + (QuoteId);
    }
    
    W = 400;
    H = 280;
    L = (screen.availWidth - W) / 2;
    T = (screen.availHeight - H) / 2;

    DetailWinParams = 'toolbar=no,location=no,status=no,menubar=no,copyhistory=no';
    DetailWinParams += ',scrollbars=no,statusbar=no,resizable=yes';
    DetailWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=' + T +'';
    
    DetailURL = RedoWebLocation
    DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);

    // Local Home & my pc at work
    // RedoWebLocation = 'Redo.asp?QuoteId=' + (QuoteId);

    // All work PC's with & without (RL.asp), If use MC-SQL2012 PC
    // Seems to be ok without as its the default document
    // RedoWebLocation = 'http://192.168.20.205/Redo/?QuoteId=' + (QuoteId);
    // RedoWebLocation = 'http://192.168.20.205/Redo/RL.asp?QuoteId=' + (QuoteId);

    // If ever need external link with & without (Redo.asp), If use MC-SQL2012 PC
    // RedoWebLocation = 'http://81.133.64.19/Redo/?QuoteId=' + (QuoteId);
    // RedoWebLocation = 'http://81.133.64.19/Redo/RL.asp?QuoteId=' + (QuoteId);

    // If Use Qc on File-Server have you have to have (RL.asp)
    // RedoWebLocation = 'http://81.133.64.20/Qc/RL.asp?QuoteId=' + (QuoteId);
    // If Use own web site on File-Server uses default document
    // RedoWebLocation = 'http://81.133.64.20/Redo/?QuoteId=' + (QuoteId);
    
    // RL.asp is used to load a swirling Gif, it then relocates to Redo.asp
    
}

function ClarityQc(JobNumber) 
{

    var DetailWinParams, DetailURL, DetailWindow, W, H, L;


    W = screen.availWidth - 200;
    H = screen.availHeight - 150;
    L = (screen.availWidth - W) / 2;

    DetailWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    DetailWinParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    DetailWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=40';
    DetailURL = 'ClarityQc.asp?Job=' + (JobNumber);
    DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);

}

function ClarityShowdetail(JobNumber)
{
    
    var DetailWinParams, DetailURL, DetailWindow, W, H, L;
        
    W = screen.availWidth - 200;
    H = screen.availHeight - 150;
    L = (screen.availWidth - W) / 2;
            
	DetailWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
	DetailWinParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
	DetailWinParams += ',width='+ W + ',height=' + H + ',left=' + L + ',top=40';
	DetailURL = 'ClarityShowDetail.asp?Job='+ (JobNumber);
	DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);

}

function ClaritySubDetail(ItemId)
 {

    var DetailWinParams, DetailURL, DetailWindow, W, H, L;

    W = screen.availWidth - 400;
    H = screen.availHeight - 250;
    L = (screen.availWidth - W) / 2;

    DetailWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
    DetailWinParams += ',scrollbars=yes,statusbar=yes,resizable=yes';
    DetailWinParams += ',width=' + W + ',height=' + H + ',left=' + L + ',top=40';
    DetailURL = 'ClaritySubDetail.asp?Id=' + (ItemId);
    DetailWindow = window.open(DetailURL, '_blank', DetailWinParams);

}

function ClickMsg()
{
    window.alert("Clicking the boxes will not update the Database!")
}

function ShowDetailLoadChk()
{
   // Window will close after 5 minutes
   // Works on detail & sub detail
   setTimeout("window.close()", 300000);
}

function LogonLoadChk() 
{
    var focusHere = document.frmLogon.txtUserID;
    focusHere = focusHere.focus();
    return ErrChk();
}

function ErrChk() 
{
    var emsg = document.getElementById("ErrBox").value
    if (document.getElementById("ErrBox").value.length > 1)
    { window.alert(emsg); }
}

//functions for AddUser

function AddUsersFormReset()
{
    getObject('hAction').value = 'reset';
    getObject('frmAddUser').submit();
}

function AddUsersAddRecord()
{
    getObject('hAction').value = 'Add';
    getObject('frmAddUser').submit();
}

function Export(Address)
{

   // var ExportWinParams, ExportURL, ExportWindow, W, H, L; 

   // var ExportWinParams, ExportURL, ExportWindow;
   // var MyTop = (screen.availHeight / 2) - 78;
   // var MyLeft = (screen.availWidth / 2) - 125;
   // ExportWinParams = 'toolbar=no,location=no,status=yes,menubar=no,copyhistory=no';
   // ExportWinParams += ',statusbar=yes,resizable=yes,width=250,height=155,left=' + MyLeft + ',top=' + MyTop;   //left=100,top=100';
   // ExportURL = 'http://127.0.0.1/LabourExport/ExportMsg.asp'
   // ExportWindow = window.open(ExportURL, '_blank', ExportWinParams);
    
    window.location.replace(Address);

}


