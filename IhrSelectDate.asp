<%@language="vbscript" codepage="1252"%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

Dim ErrMsg
ErrMsg = ""

ErrMsg = Session("JobNoError")

Session("JobNo") = ""
Session("JobNoError") = ""

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>In House Redo</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/IhRedoJSFunc.js"></script>
    <link rel="shortcut icon" href="Images/X1.ico" type="image/x-icon" /> 
    <script type="text/javascript">
        function formSubmit(Selection) 
        {
            document.body.style.cursor = 'wait';
            document.getElementById('loading').innerHTML = "Loading data, please wait...";
            document.getElementById("Options").style.opacity = 0;
            document.getElementById("GoBack").style.color = "#ffffff";
            if (Selection == '0')
            { window.location.replace("IhrDisplay.asp"); }
            else
            { window.location.replace("IhrDisplayByDate.asp?sd=" + Selection); }
        }

        function formSubmitMyLogs() 
        {
            document.body.style.cursor = 'wait';
            document.getElementById('loading').innerHTML = "Loading data, please wait...";
            document.getElementById("Options").style.opacity = 0;
            document.getElementById("GoBack").style.color = "#ffffff";
            window.location.replace("IhrDisplayById.asp?sd=5");
        }

        function EchkLocal() 
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
        
    </script>
</head>

<body style="padding: 0px; margin: 0px" onload="EchkLocal();">


<table  style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
		<td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>            
    </tr>
	<tr>
		<td height="8" valign="top" colspan="3">
		    <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		</td>
	</tr>
	 
	<tr>
	    <td  style="height: 20px" width="33%">
			&nbsp;<a href ="javascript:window.location.replace('IhrJobNo.asp');" id="GoBack" style="font-size:12px; color: <%=NewBlue%>;">Return to main page</a>
		</td>
	    <td align="center" valign="bottom" width="34%">
	        <img align="top" alt="mediaco logo" src="Images/X1.ico" style="width: 20px; height: 20px;" />&nbsp;
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">&nbsp;Select Date</font>                     
        </td>
        <td align="right" valign="bottom" width="33%">
            <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;
        </td>
	</tr>	
</table>

<form action="IhrCheckViewJobNo.asp" method="post" name="frmViewJobNo" id="frmViewJobNo" onsubmit="return ValidateRedoJobNo();">

<table id="Options" style="width: 100%; position: absolute; top: 40%; padding-right: 10px; padding-left: 10px;" >
    <tr>   
        <td >
			<font size="3" style="color:<%=NewBlue%>;">
			    
				
			    &nbsp;<input name="RadioAll" id="RadioAll" type="radio" value="ON" onclick="javascript:formSubmit('0');" />&nbsp;All
			    <br /><br />
			    &nbsp;<input name="RadioToday" id="RadioToday" type="radio" value="ON" onclick="javascript:formSubmit('6');" />&nbsp;Today
			    <br /><br />
                &nbsp;<input name="RadioThisWeek" id="RadioThisWeek" type="radio" value="ON" onclick="javascript:formSubmit('2');" />&nbsp;This&nbsp;Week
			    <br /><br />
			    &nbsp;<input name="RadioThisMonth" id="RadioThisMonth" type="radio" value="ON" onclick="javascript:formSubmit('3');" />&nbsp;This&nbsp;Month
			    <br /><br />
			    &nbsp;<input name="RadioThisYear" id="RadioThisYear" type="radio" value="ON" onclick="javascript:formSubmit('4');" />&nbsp;This&nbsp;Year
			    <br /><br />
			    <%
			    If Session("AddRedo") = Cbool(True) Then			    
			    %>
			    &nbsp;<input name="RadioMyLogs" id="RadioMyLogs" type="radio" value="ON" onclick="javascript:formSubmitMyLogs();" />&nbsp;My&nbsp;logs
			    <br /><br />
			    <%End If%>
			    &nbsp;&nbsp;<input name="txtJobNo" id="txtJobNo" type="text" value="" />&nbsp;By&nbsp;Job
			    &nbsp;&nbsp;<input name="ErrBox" id="ErrBox" type="hidden" value="<%=ErrMsg%>" />
			</font>				
		</td>
    </tr>          
</table>

</form>
      
<table style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;" >
    <tr>  
        <td height="70">
            <hr style="border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
            Trafford Park, Manchester M17 1BS <br />Tel:(+44)161 875 2020 Fax:(+44)161 873 7740</font></p>
        </td>
    </tr>
</table>
<span id="loading" style="font-size: medium; color:<%=NewBlue%>; font-weight: bold; position: absolute; top: 50%; left: 43%;"></span>
<!--</form>-->
</body>  
</html>