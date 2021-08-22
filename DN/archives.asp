<% option explicit %>
<Head>
<title>DailyNews Archives</title>
</head>
<body>
<!--p align="center">


</p-->
<!--p align="center"><span lang="en-us">Click on the date to retrieve the issue of that date. </span-->
<center>
<%
	Dim anclObj, myDate
	If Request.Form("anclYear") <> "" Then
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