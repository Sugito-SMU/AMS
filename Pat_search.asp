<%@ Language=VBScript %>
<!-- #INCLUDE FILE="../includes/CheckParams.asp" -->
<%
	system = "SYSPATRON"
	code   = "PATRONUPDATE"
	access = "N"
%>
<!-- #INCLUDE FILE="secure.asp" -->
<!-- #INCLUDE FILE="AdoVbs.Inc"  -->


<%
Function LeadingZero(str)
      LeadingZero = Right("00" & str,2)
End Function
   
Function FormatDateToDDMMYYYY(ClockDate)		' ClockDate in MM/DD/YYYY format, return in DD/MM/YYYY format
	IF (ClockDate <> "") THEN
		FormatDateToDDMMYYYY = LeadingZero(Day(ClockDate)) & "/" & LeadingZero(Month(ClockDate)) & "/" & Year(ClockDate)
	ELSE
		FormatDateToDDMMYYYY = ""
	END	IF
End Function

IF REQUEST.FORM("submitsearch")<>"" THEN
	NetworkId=LTRIM(ValidateAndEncodeXSSEx(REQUEST.FORM("NetworkId")))
	PatName=LTRIM(ValidateAndEncodeXSSEx(REQUEST.FORM("PatName")))
	PatDept=LTRIM(ValidateAndEncodeXSSEx(REQUEST.FORM("PatDept")))
	SortCon=LTRIM(ValidateAndEncodeXSSEx(REQUEST.FORM("SortCon")))
	ActiveFlg=LTRIM(ValidateAndEncodeXSSEx(REQUEST.FORM("ActiveFlg")))
END IF
%>
<%
SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
SET CurrCmd.ACTIVECONNECTION = OBJCONN
strSQL = "SELECT * FROM V_DEPARTMENT ORDER BY DEPT_T"
CurrCmd.CommandText = strSQL
CurrCmd.CommandType = adCmdText
set DeptList = CurrCmd.EXECUTE()
%>

<!-- #INCLUDE FILE="header.asp" -->
<br>
<table ID="searchtable1" width="90%" align="center"  >
<form method="POST" action="pat_search.asp" name="searchform">
<tr>
	<td  align="center" valign="middle" class="searchbar"><b>
	Network ID: <input name="networkid" value="<%=ValidateAndEncodeXSSEx(NetworkId)%>" type="text" size="12" maxlength="12" >
	&nbsp;Name: <input name="patname" value="<%=ValidateAndEncodeXSSEx(PatName)%>" type="text" size="20" maxlength="20" class="smalllink">
	&nbsp;Dept:
	<select name="patdept">
		<option value="">
		<%do while not DeptList.eof%>
			<option value="<%=ValidateAndEncodeXSSEx(DeptList("DEPT_C"))%>" <%IF PatDept=UCASE(DeptList("DEPT_C")) THEN%>selected<%END IF%>><%=ValidateAndEncodeXSSEx(DeptList("DEPT_T"))%> [<%=ValidateAndEncodeXSSEx(DeptList("DEPT_C"))%>]
		<%
			DeptList.movenext
			loop
		%>
	</select>
	<hr width="100%" color="black" size=1>
	Patron Status:
	<select name="activeflg">
		<option>
		<option value="Y" <%IF ActiveFlg="Y" THEN%>selected<%END IF%>>Active
		<option value="N" <%IF ActiveFlg="N" THEN%>selected<%END IF%>>Inactive
		<option value="F" <%IF ActiveFlg="F" THEN%>selected<%END IF%>>Inactive (Future-dated)
	</select>

	&nbsp;
	<input type="submit" value="Search" name="submitsearch">
	</b>
</td></tr>
</form>
</table>
<br><br>

<%IF TRIM(NetworkId)<>"" OR  TRIM(PatName)<>"" OR  TRIM(PatDept)<>"" OR  TRIM(ActiveFlg)<>"" THEN%>

<%
'******************************
'*****Request Variables *******
'******************************
SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
SET CurrCmd.ACTIVECONNECTION = OBJCONN
strSQL = " SELECT ACTIVE_NUM,LOGON_T,NAME_T,DEPT_C,PATRON_CD,ACTIVE_F,PATRON_T, DEPT_T,FACULTY_C,FACULTY_T, WORK_START_DT, WORK_END_DT, FDT_REC_F "&_
		" FROM V_PATRON_PRIDTL WHERE (?='' OR CHARINDEX(?,LOGON_T)>0) AND (?='' OR CHARINDEX(?,NAME_T)>0) AND (?='' OR DEPT_C=?)AND (?='' OR ACTIVE_F=?) "&_
 		" UNION "&_
		" SELECT ACTIVE_NUM,LOGON_T,NAME_T,DEPT_C,PATRON_CD,ACTIVE_F,PATRON_T, DEPT_T,FACULTY_C,FACULTY_T, WORK_START_DT, WORK_END_DT, FDT_REC_F "&_
		" FROM V_PATRON_PRIDTL_FDT WHERE (?='' OR CHARINDEX(?,LOGON_T)>0) AND (?='' OR CHARINDEX(?,NAME_T)>0) AND (?='' OR DEPT_C=?)AND (?='' OR ACTIVE_F=? OR ?='F') "&_
		" ORDER BY NAME_T, WORK_START_DT"
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("NetworkId1", adVarChar,adParamInput,50,ValidateAndEncodeSQL(NetworkId))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("NetworkId2", adVarChar,adParamInput,50,ValidateAndEncodeSQL(NetworkId))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatName1", adVarChar,adParamInput,50,ValidateAndEncodeSQL(PatName))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatName2", adVarChar,adParamInput,50,ValidateAndEncodeSQL(PatName))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatDept1", adVarChar,adParamInput,10,ValidateAndEncodeSQL(PatDept))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatDept2", adVarChar,adParamInput,10,ValidateAndEncodeSQL(PatDept))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ActiveFlg1", adVarChar,adParamInput,10,ValidateAndEncodeSQL(ActiveFlg))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ActiveFlg2", adVarChar,adParamInput,10,ValidateAndEncodeSQL(ActiveFlg))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("NetworkId3", adVarChar,adParamInput,50,ValidateAndEncodeSQL(NetworkId))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("NetworkId4", adVarChar,adParamInput,50,ValidateAndEncodeSQL(NetworkId))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatName3", adVarChar,adParamInput,50,ValidateAndEncodeSQL(PatName))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatName4", adVarChar,adParamInput,50,ValidateAndEncodeSQL(PatName))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatDept3", adVarChar,adParamInput,10,ValidateAndEncodeSQL(PatDept))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PatDept4", adVarChar,adParamInput,10,ValidateAndEncodeSQL(PatDept))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ActiveFlg3", adVarChar,adParamInput,10,ValidateAndEncodeSQL(ActiveFlg))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ActiveFlg4", adVarChar,adParamInput,10,ValidateAndEncodeSQL(ActiveFlg))
CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ActiveFlg5", adVarChar,adParamInput,10,ValidateAndEncodeSQL(ActiveFlg))
CurrCmd.CommandText = strSQL
CurrCmd.CommandType = adCmdText
set PatronList = CurrCmd.EXECUTE()
	
%>
	<%If not PatronList.EOF then%>
	<table width="90%" border=0 nowrap align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
	<tr bgcolor="#FF9900"> 
		<td style="text-align:center;"><b>Patron#</b></td> 
		<td><b>Name</b></td> 
		<td><b>Network ID</b></td> 
		<td><b>Patron Type</b></td> 
		<td><b>Faculty Type</b></td> 
		<td><b>Primary Department</b></td>
		<td style="text-align:center;"><b>Access Start Date</b></td>
		<td style="text-align:center;"><b>Access End Date</b></td>
		<td style="text-align:center;"><b>Status</b></td> 
	</tr> 
	<%
		r_count = 0
		while not PatronList.EOF
			r_count = r_count + 1
	%>
	<tr <%IF r_count MOD 2=0 THEN%>bgcolor="#eeeedd"<%ELSE%>bgcolor="#FFFFFF"<%END IF%>>
		<td style="text-align:center;"><a href="Pat_Upd.asp?ActiveNum=<%=ValidateAndEncodeXSSEx(PatronList("ACTIVE_NUM"))%>&FDated=<%=ValidateAndEncodeXSSEx(PatronList("FDT_REC_F"))%>"><%=PatronList("ACTIVE_NUM")%></a></td> 
		<td><a href="Pat_Upd.asp?ActiveNum=<%=ValidateAndEncodeXSSEx(PatronList("ACTIVE_NUM"))%>&FDated=<%=ValidateAndEncodeXSSEx(PatronList("FDT_REC_F"))%>"><%=ValidateAndEncodeXSSEx(PatronList("NAME_T"))%></a></td> 
		<td><%=ValidateAndEncodeXSSEx(PatronList("LOGON_T"))%></td> 
		<td><%=ValidateAndEncodeXSSEx(PatronList("PATRON_T"))%></td> 
		<td><%=ValidateAndEncodeXSSEx(PatronList("Faculty_T"))%></td> 
		<td><%=ValidateAndEncodeXSSEx(PatronList("DEPT_T"))%></td> 
		<td style="text-align:center;"><%=ValidateAndEncodeXSSEx(FormatDateToDDMMYYYY(PatronList("WORK_START_DT")))%></td> 
		<td style="text-align:center;"><%=ValidateAndEncodeXSSEx(FormatDateToDDMMYYYY(PatronList("WORK_END_DT")))%></td> 
		<td style="text-align:center;">
			<%if PatronList("ACTIVE_F") = "Y" then%>
				<span>Active</span>
			<%else%>
				<span class="response_neg">Inactive</span>
			<%end if%>
			<%IF PatronList("FDT_REC_F") = "Y" THEN %>
				<br>
				<span style="color:red;font-weight:bold;">(Future-dated)</span>
			<%END IF%>
		</td> 
	</tr>
	<%
		PatronList.MoveNext
		wend%>
	</table>
	<%else%>
		<div align="center">No records found.</div>
	<%end if%>
<%ELSE%>
	<div align="center">Please input a search criteria.</div>
<%END IF%>
<bR><br>
<!-- #INCLUDE FILE="footer.asp" -->
