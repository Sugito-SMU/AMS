<%'<!-- #INCLUDE FILE="../checkparams.asp"  -->%>
<%VersionNum=1.01%>
<html>
<head>
	<title>Account Management System :: Singapore Management University</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr> 
	<td width="100%"  style="background-color:#FFAB08;font-size:18pt; color:white; border: 1px solid #454646" align="left">
Accounts Management (Patron)</td>
  </tr>
  <tr bgcolor="#000000"> 
    <td colspan="2" class="functionbar">
		&nbsp;&nbsp;
		<a href="pat_search.asp"><span style="text-decoration:none;<%if Functiontype="patsearch" THEN%>color:pink;font-weight:bold;<%ELSE%>color:white;<%END IF%>">Search Record</a></span>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="pat_upd.asp"><span style="text-decoration:none;<%if Functiontype="patadd" THEN%>color:pink;font-weight:bold;<%ELSE%>color:white;<%END IF%>">Add Record</a></span>
		&nbsp;&nbsp;
</td>
  </tr>
	
</table>



