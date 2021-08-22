<% option explicit 
on error  resume next
dim id1
id1=Request.QueryString("id")%>

<Head>
<title><% If id1 = 0 Then  
		Response.write("Dinamina Archives")
		else 
		Response.write("Silumina Archives")
		end if %></title>
</head>
<body onload="self.resizeTo(250,450); self.moveTo(((screen.width/2)-125),((screen.height/2)-200)); return false;">
<%	If id1 = 1 Then %>
	<!--img border="0" src="/img/ArchSilu.jpg" width="182" height="50"-->
	<% Else %>
	<!--img border="0" src="/img/ArchDina.jpg" width="182" height="50"-->
	<% End If %>
	<!--img border="0" src="/img/archives_.jpg" width="103" height="23"><p-->
<!--span lang="en-us">&#3476;&#3510;&#3495; &#3461;&#3520;&#3521;&#3530;&#8205;&#3514; &#3482;&#3517;&#3535;&#3508;&#3514; &#3501;&#3549;&#3515;&#3535; &#3484;&#3536;&#3505;&#3539;&#3512;&#3495; &#3461;&#3503;&#3535;&#3517; &#3503;&#3538;&#3505;&#3514; &#3512;&#3501; &#3482;&#3530;&#3517;&#3538;&#3482;&#3530; &#3482;&#3515;&#3505;&#3530;&#3505;. </span-->
<center>
<%
	Dim anclObj, myDate
	If (Request.Form("anclYear") <> "") Then
		myDate = DateSerial(Request.Form("anclYear"), Request.Form("anclMonth"), Request.Form("anclDay"))
	End If
	
	Set anclObj = GetObject("script:" & Server.MapPath(".\anclCalendar.wsc"))

	If isDate(myDate) Then anclObj.anclDate = myDate
	With anclObj
		.anclActionURL = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING") 
	End With
	anclObj.displayCalendar()
	Set anclObj = Nothing
%>
</center>
</body>