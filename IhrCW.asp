
<% 
' *** End Session & close window 
Response.Cookies("IpChecked").expires = Date -5
Session.Abandon 
%>
    
<script language="JavaScript" type="text/javascript">
	{window.close();} 
</script>
			
