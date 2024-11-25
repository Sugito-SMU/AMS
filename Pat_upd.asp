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

'Start -- Added by Avy - 6 Jul 09
Function LeadingZero(str)
      LeadingZero = Right("00" & str,2)
End Function
   
Function FormatDateToDDMMYYYY(ClockDate)		' ClockDate in MM/DD/YYYY format, return in DD/MM/YYYY format
      FormatDateToDDMMYYYY = LeadingZero(Day(ClockDate)) & "/" & LeadingZero(Month(ClockDate)) & "/" & Year(ClockDate)
End Function

'End -- Added by Avy - 6 Jul 09

Function FormatMMM(mm)
	dim mmm
	IF (CINT(mm) = 1) THEN
		mmm = "Jan"
	ELSEIF (CINT(mm) = 2) THEN
		mmm = "Feb"
	ELSEIF (CINT(mm) = 3) THEN
		mmm = "Mar"
	ELSEIF (CINT(mm) = 4) THEN
		mmm = "Apr"
	ELSEIF (CINT(mm) = 5) THEN
		mmm = "May"
	ELSEIF (CINT(mm) = 6) THEN
		mmm = "Jun"
	ELSEIF (CINT(mm) = 7) THEN
		mmm = "Jul"
	ELSEIF (CINT(mm) = 8) THEN
		mmm = "Aug"
	ELSEIF (CINT(mm) = 9) THEN
		mmm = "Sep"
	ELSEIF (CINT(mm) = 10) THEN
		mmm = "Oct"
	ELSEIF (CINT(mm) = 11) THEN
		mmm = "Nov"
	ELSEIF (CINT(mm) = 12) THEN
		mmm = "Dec"
	END IF
	FormatMMM = mmm
End Function



'******************************
'*****Request Variables *******
'******************************
IF Request.querystring("ActiveNum") <> "" THEN
	v_ActiveNum = CInt(Request.querystring("ActiveNum"))
ELSE
	v_ActiveNum = 0
END IF
%>

<%
IF REQUEST.FORM("UpdateSubmit")<>"" OR REQUEST.FORM("AddSubmit")<>"" THEN
	IF v_ActiveNum<>0 THEN
		'******************************
		SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
		SET CurrCmd.ACTIVECONNECTION = OBJCONN
		IF (REQUEST.QUERYSTRING("FDated")="Y") THEN
			strSQL = " SELECT ACTIVE_NUM, LOGON_T, NAME_T, FIRST_NM, LAST_NM, EMAIL_T, SMU_TEL_N, PATRON_CD, FACULTY_C, UPDATE_F, SALUTATION_C, SALUTATION_T, EMP_N, DEPT_C, FACULTY_T, WORK_START_DT, WORK_END_DT, ACTIVE_F, CREATION_ID, CONVERT(VARCHAR, CREATION_DT, 113) AS CREATION_DT, LAST_UPD_ID, CONVERT(VARCHAR, LAST_UPD_DT, 113) AS LAST_UPD_DT, PATRON_T, DEPT_T, TITLE, DUMMY_F, OFFICIAL_NM, CV_URL, FDT_REC_F, HAS_FDT_F "&_
					" FROM V_PATRON_PRIDTL_FDT WHERE ACTIVE_NUM=?"
		ELSE
			strSQL = " SELECT ACTIVE_NUM, LOGON_T, NAME_T, FIRST_NM, LAST_NM, EMAIL_T, SMU_TEL_N, PATRON_CD, FACULTY_C, UPDATE_F, SALUTATION_C, SALUTATION_T, EMP_N, DEPT_C, FACULTY_T, WORK_START_DT, WORK_END_DT, ACTIVE_F, CREATION_ID, CONVERT(VARCHAR, CREATION_DT, 113) AS CREATION_DT, LAST_UPD_ID, CONVERT(VARCHAR, LAST_UPD_DT, 113) AS LAST_UPD_DT, PATRON_T, DEPT_T, TITLE, DUMMY_F, OFFICIAL_NM, CV_URL, FDT_REC_F, HAS_FDT_F "&_
					" FROM V_PATRON_PRIDTL WHERE ACTIVE_NUM=?"
		END IF

		CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("v_ActiveNum", adInteger,adParamInput,,ValidateAndEncodeSQL(v_ActiveNum))
		CurrCmd.CommandText = strSQL
		CurrCmd.CommandType = adCmdText
		set PatroNList = CurrCmd.EXECUTE()

		if PatronList.EOF then
			response.write "Record not found.  Click <a href='pat_search.asp'>here</a> to go back to the patron list."
			response.end
		end if
		if PatronList("Update_f") = "N" THEN
			response.write "Update Not Allowed"
			response.end
		end if
		ActiveNum_Val=CInt(PatronList("ACtive_Num"))
	ELSE
		ActiveNum_Val=0
	END IF
			
	'Start -- Updated by Avy - 6 Jul 09
	IF TRIM(REQUEST.FORM("FromDte"))="" THEN
		V_StartDte=NULL
	ELSE
		V_StartDte = CDate(mid(TRIM(REQUEST.FORM("FromDte")),1,2) & "-" & FormatMMM(mid(TRIM(REQUEST.FORM("FromDte")),4,2)) & "-" & mid(TRIM(REQUEST.FORM("FromDte")),7,4)&" 00:00:00")
	END IF
	
	
	IF TRIM(REQUEST.FORM("ToDte"))="" THEN
		V_EndDte=NULL
	ELSE
		V_EndDte=CDate(mid(TRIM(REQUEST.FORM("ToDte")),1,2) & "-"  & FormatMMM(mid(TRIM(REQUEST.FORM("ToDte")),4,2)) & "-" & mid(TRIM(REQUEST.FORM("ToDte")),7,4)&" 23:59:59")
	END IF
	
	'End -- Updated by Avy - 6 Jul 09	
	
	IF TRIM(Request.form("SMUTEL"))="" THEN
		V_SMU_TEL=NULL	
	ELSE
		V_SMU_TEL=TRIM(Request.form("SMUTEL"))
	END IF
	IF TRIM(Request.form("EMpNum"))="" THEN
		V_EMP_N=NULL	
	ELSE
		V_EMP_N=TRIM(Request.form("EMpNum"))
	END IF
	IF V_DUMMY_DEF="I" THEN
		V_DUMMY_F=TRIM(Request.form("DUMMY_F"))
	ELSE
		V_DUMMY_F=V_DUMMY_DEF
	END IF
	IF TRIM(Request.form("PriDeptCd"))="" THEN
	V_PRI_DEPT_C=""
	ELSE
		V_PRI_DEPT_C=MID(Request.form("PriDeptCd"),1,INSTR(Request.form("PriDeptCd"),"|")-1)
	END IF
	IF TRIM(Request.form("PriTitle"))="" THEN 
		V_PRI_TITLE=""	
	ELSE
		V_PRI_TITLE=TRIM(Request.form("PriTitle"))
	END IF
	IF TRIM(Request.form("PATRONTypeCD"))="" THEN 
	V_PATRON_CD=""
	ELSE
		V_PATRON_CD=MID(Request.form("PATRONTypeCD"),1,INSTR(Request.form("PATRONTypeCD"),"|")-1)
	END IF
	IF TRIM(Request.form("FACTypeCD"))="" THEN
	V_FACULTY_C=""
	ELSE
		V_FACULTY_C=MID(Request.form("FACTypeCD"),1,INSTR(Request.form("FACTypeCD"),"|")-1)
	END IF
	V_SEC_DEPT_1=Request.form("SecDeptCd1")
	V_SEC_DEPT_2=Request.form("SecDeptCd2")
	V_SEC_DEPT_3=Request.form("SecDeptCd3")
	IF TRIM(Request.form("SecTitle1"))="" THEN
		V_SEC_TITLE_1=NULL	
	ELSE
		V_SEC_TITLE_1=TRIM(Request.form("SecTitle1"))
	END IF
	IF TRIM(Request.form("SecTitle2"))="" THEN
		V_SEC_TITLE_2=NULL	
	ELSE
		V_SEC_TITLE_2=TRIM(Request.form("SecTitle2"))
	END IF
	IF TRIM(Request.form("SecTitle3"))="" THEN
		V_SEC_TITLE_3=NULL	
	ELSE
		V_SEC_TITLE_3=TRIM(Request.form("SecTitle3"))
	END IF
	V_NAME_T=TRIM(Request.form("DisplayNm"))
	V_OFF_NAME_T=TRIM(Request.form("OfficialNm"))
	V_FIRST_NM=TRIM(Request.form("FirstNm"))
	V_LAST_NM=TRIM(Request.form("LASTNM"))
	V_SALUTATION_C=Request.form("SaluteCd")
	V_CV_URL=TRIM(Request.form("CV_URL"))
	V_LOGON_T=TRIM(Request.form("LOGON_T"))
	V_EMAIL_T=TRIM(Request.form("EMAIL_T"))

	SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
	SET CurrCmd.ACTIVECONNECTION = OBJCONN
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("Active_Num", adBigInt,adParamInput,,ActiveNum_Val)
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("Name_t", adVarchar,adParamInput,50,DecodeDoubleQuotes(V_NAME_T))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("OFF_NAME_T", adVarchar,adParamInput,200,DecodeDoubleQuotes(V_OFF_NAME_T))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("first_nm", adVarchar,adParamInput,80,DecodeDoubleQuotes(V_FIRST_NM))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("last_nm", adVarchar,adParamInput,25,DecodeDoubleQuotes(V_LAST_NM))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("patron_cd", adVarchar,adParamInput,2,DecodeDoubleQuotes(V_PATRON_CD))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("faculty_c", adVarchar,adParamInput,2,DecodeDoubleQuotes(V_FACULTY_C))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("EMP_N", adVarchar,adParamInput,10,DecodeDoubleQuotes(V_EMP_N))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("START_DT", adDBTimeStamp,adParamInput,,(V_STARTDTE))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("END_DT", adDBTimeStamp,adParamInput,,(V_ENDDTE))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SALUTATION_C", adVarchar,adParamInput,10,DecodeDoubleQuotes(V_SALUTATION_C))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SMU_TEL", adVarchar,adParamInput,15,DecodeDoubleQuotes(V_SMU_TEL))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PriDept_c", adVarchar,adParamInput,10,DecodeDoubleQuotes(V_PRI_DEPT_C))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("PriDept_title", adVarchar,adParamInput,200,DecodeDoubleQuotes(V_PRI_TITLE))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecDept_1", adVarchar,adParamInput,20,DecodeDoubleQuotes(V_SEC_DEPT_1))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecDept_2", adVarchar,adParamInput,20,DecodeDoubleQuotes(V_SEC_DEPT_2))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecDept_3", adVarchar,adParamInput,20,DecodeDoubleQuotes(V_SEC_DEPT_3))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecTitle_1", adVarchar,adParamInput,200,DecodeDoubleQuotes(V_SEC_TITLE_1))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecTitle_2", adVarchar,adParamInput,200,DecodeDoubleQuotes(V_SEC_TITLE_2))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("SecTitle_3", adVarchar,adParamInput,200,DecodeDoubleQuotes(V_SEC_TITLE_3))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("CV_URL", adVarchar,adParamInput,500,DecodeDoubleQuotes(V_CV_URL))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("LOGON_T", adVarchar,adParamInput,30,DecodeDoubleQuotes(V_LOGON_T))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("EMAIL_T", adVarchar,adParamInput,40,DecodeDoubleQuotes(V_EMAIL_T))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("DUMMY_F", adVarchar,adParamInput,4,DecodeDoubleQuotes(V_DUMMY_F))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("CurrUser", adVarchar,adParamInput,50,DecodeDoubleQuotes(Ucase(Trim(Request.ServerVariables("LOGON_USER")))))
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ResCSP", adBigInt, adParamOutput)
	CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("NewActiveNum", adBigInt, adParamOutput)

	fDatedFlg = false
	IF (v_ActiveNum <> 0) THEN
		IF (( ( ((PatronList("FACULTY_C") <> V_FACULTY_C) AND (V_FACULTY_C = "AJ")) OR _
				((PatronList("PATRON_CD") <> V_PATRON_CD) AND (V_PATRON_CD = "03")) ) _
				AND (V_STARTDTE > Date) AND ((PatronList("ACTIVE_F") = "Y"))) _
			OR (REQUEST.QUERYSTRING("FDated")="Y")) THEN
			CurrCmd.COMMANDTEXT = "CSP_PATRON_DTL_FDT_UPD"
			fDatedFlg = true

		ELSE
			CurrCmd.COMMANDTEXT = "CSP_PATRON_DTL_UPD"
		END IF	
	ELSE
		CurrCmd.COMMANDTEXT = "CSP_PATRON_DTL_UPD"
	END IF

	CurrCmd.EXECUTE ,,adCmdStoredProc

	IF CINT(CurrCmd.Parameters("ResCSP").value)=0 Then
		IF (ActiveNum_Val > 0) THEN
			newActiveNum = ActiveNum_Val
		ELSE
			newActiveNum = CurrCmd.Parameters("NewActiveNum").value
		END IF
		respUrl = "pat_upd.asp?uflg=1&activenum="&newActiveNum
		IF (fdatedFlg = true) OR (REQUEST.QUERYSTRING("FDated")="Y") THEN
			respUrl = respUrl & "&FDated=Y"
		END IF
		RESPONSE.REDIRECT respUrl
	ELSEIF CINT(CurrCmd.Parameters("ResCSP").value)=1 Then
		SubmitError="Invalid Patron/Faculty code."
	ELSEIF CINT(CurrCmd.Parameters("ResCSP").value)=3 Then
		SubmitError="Same Network ID exists for another patron record."
	ELSEIF CINT(CurrCmd.Parameters("ResCSP").value)=4 Then
		SubmitError="Primary department already exists."
	ELSE
		SubmitError="Error ["&CurrCmd.Parameters("ResCSP").value&"] - Contact Administrator"
	END IF

ELSEIF REQUEST.FORM("DeleteSubmit")<>"" THEN
	IF v_ActiveNum<>0 THEN
		SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
		SET CurrCmd.ACTIVECONNECTION = OBJCONN
		CurrCmd.COMMANDTEXT = "CSP_PATRON_DTL_FDT_DEL"
		CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("Active_Num", adBigInt,adParamInput,,v_ActiveNum)
		CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("ResCSP", adBigInt, adParamOutput)
		CurrCmd.EXECUTE ,,adCmdStoredProc
		IF CINT(CurrCmd.Parameters("ResCSP").value)=0 THEN
			respUrl = "pat_upd.asp?uflg=2&activenum="&v_ActiveNum
			RESPONSE.REDIRECT respUrl
		ELSE
			SubmitError="Error ["&CurrCmd.Parameters("ResCSP").value&"] - Contact Administrator"
		END IF
	END IF
END IF

'----- CATEGORIZE
		IF v_ActiveNum<>0 THEN
			UpdateMode="Update"
			'******************************
			SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
			SET CurrCmd.ACTIVECONNECTION = OBJCONN
			IF (REQUEST.QUERYSTRING("FDated")="Y") THEN
				strSQL = "SELECT ACTIVE_NUM, LOGON_T, NAME_T, FIRST_NM, LAST_NM, EMAIL_T, SMU_TEL_N, PATRON_CD, FACULTY_C, UPDATE_F, SALUTATION_C, SALUTATION_T, EMP_N, DEPT_C, FACULTY_T, WORK_START_DT, WORK_END_DT, ACTIVE_F, CREATION_ID, CONVERT(VARCHAR, CREATION_DT, 113) AS CREATION_DT, LAST_UPD_ID, CONVERT(VARCHAR, LAST_UPD_DT, 113) AS LAST_UPD_DT, PATRON_T, DEPT_T, TITLE, DUMMY_F, OFFICIAL_NM, CV_URL, FDT_REC_F, HAS_FDT_F "&_
					" FROM V_PATRON_PRIDTL_FDT a "&_
					" WHERE a.ACTIVE_NUM = ?"
			ELSE
				strSQL = "SELECT ACTIVE_NUM, LOGON_T, NAME_T, FIRST_NM, LAST_NM, EMAIL_T, SMU_TEL_N, PATRON_CD, FACULTY_C, UPDATE_F, SALUTATION_C, SALUTATION_T, EMP_N, DEPT_C, FACULTY_T, WORK_START_DT, WORK_END_DT, ACTIVE_F, CREATION_ID, CONVERT(VARCHAR, CREATION_DT, 113) AS CREATION_DT, LAST_UPD_ID, CONVERT(VARCHAR, LAST_UPD_DT, 113) AS LAST_UPD_DT, PATRON_T, DEPT_T, TITLE, DUMMY_F, OFFICIAL_NM, CV_URL, FDT_REC_F, HAS_FDT_F " &_
					" FROM V_PATRON_PRIDTL a "&_
					" WHERE a.ACTIVE_NUM = ?"
			END IF
			CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("v_ActiveNum", adInteger,adParamInput,,ValidateAndEncodeSQL(v_ActiveNum))
			CurrCmd.CommandText = strSQL
			CurrCmd.CommandType = adCmdText
			set PatronList = CurrCmd.EXECUTE()

			if PatronList.EOF then
				response.write "Record not found.  Click <a href='pat_search.asp'>here</a> to go back to the patron list."
				response.end
			end if
			if PatronList("Update_f") = "N" THEN
				response.write "Update Not Allowed"
				response.end
			end if


			V_PRI_DEPT_C=PatronList("DEPT_C")
			V_PRI_TITLE=PatronList("TITLE")

			'******************************
			SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
			SET CurrCmd.ACTIVECONNECTION = OBJCONN
			IF (REQUEST.QUERYSTRING("FDated")="Y") THEN
				strSQL = "Select DEPT_C, TITLE FROM V_PATRON_DEPTLIST_FDT WHERE active_num=? AND PRI_DEPT_FLAG='N' ORDER BY DEPT_T DESC"
			ELSE
				strSQL = "Select DEPT_C, TITLE FROM V_PATRON_DEPTLIST WHERE active_num=? AND PRI_DEPT_FLAG='N' ORDER BY DEPT_T DESC"
			END IF
			CurrCmd.PARAMETERS.APPEND CurrCmd.CreateParameter("v_ActiveNum", adInteger,adParamInput,,ValidateAndEncodeSQL(v_ActiveNum))
			CurrCmd.CommandText = strSQL
			CurrCmd.CommandType = adCmdText
			set PatronSecDeptList = CurrCmd.EXECUTE()

			IF NOT PatronSecDeptList.EOF THEN
				V_SEC_DEPT_1=PatronSecDeptList("DEPT_C")
				V_SEC_TITLE_1=PatronSecDeptList("TITLE")
				PatronSecDeptList.MoveNext
			ELSE 
				V_SEC_DEPT_1=""
				V_SEC_TITLE_1=""
			END IF
			IF NOT PatronSecDeptList.EOF THEN
				V_SEC_DEPT_2=PatronSecDeptList("DEPT_C")
				V_SEC_TITLE_2=PatronSecDeptList("TITLE")
				PatronSecDeptList.MoveNext
			ELSE 
				V_SEC_DEPT_2=""
				V_SEC_TITLE_2=""
			END IF
			IF NOT PatronSecDeptList.EOF THEN
				V_SEC_DEPT_3=PatronSecDeptList("DEPT_C")
				V_SEC_TITLE_3=PatronSecDeptList("TITLE")
				PatronSecDeptList.MoveNext
			ELSE 
				V_SEC_DEPT_3=""
				V_SEC_TITLE_3=""
			END IF
			V_PATRON_CD=PatronList("PATRON_CD")
			V_FACULTY_C=PatronList("FACULTY_C")
			V_NAME_T=PatronList("NAME_T")
			V_OFF_NAME_T=PatronList("OFFICIAL_NM")
			V_FIRST_NM=PatronList("First_Nm")
			V_LAST_NM=PatronList("Last_Nm")
			V_SALUTATION_C=PatronList("Salutation_c")
			V_SMU_TEL=PatronList("Smu_tel_n")
			V_EMP_N=PatronList("Emp_n")
			V_CV_URL=PatronList("CV_URL")
			V_LOGON_T=LCASE(PatronList("LOGON_T"))
			V_EMAIL_T=LCASE(PatronList("EMAIL_T"))
			ACCESS_START_DT=PatronList("WORK_START_DT")
			ACCESS_END_DT=PatronList("WORK_END_DT")
			V_DUMMY_F=PatronList("DUMMY_F")		
			V_HAS_FDATED_REC = PatronList("HAS_FDT_F")

			'Start -- Updated by Avy - 6 Jul 09
			IF PatronList("WORK_START_DT")<>"" THEN							
				V_ACCESS_START_DT = FormatDateToDDMMYYYY(ACCESS_START_DT)				
			ELSE							
				V_ACCESS_START_DT = ""			
			END IF
			IF PatronList("WORK_END_DT")<>"" THEN							
				V_ACCESS_END_DT = FormatDateToDDMMYYYY(ACCESS_END_DT)				
			ELSE								
				V_ACCESS_END_DT = ""				
			END IF
			'End -- Updated by Avy - 6 Jul 09
			V_EXISTING_PATRON_CD = PatronList("PATRON_CD")
			V_EXISTING_FACULTY_CD = PatronList("FACULTY_C")
	ELSE

			UpdateMode="Add"
			functiontype="patadd"
	
			V_PRI_DEPT_C=""
			V_PRI_TITLE=""
			V_SEC_DEPT_1=""
			V_SEC_DEPT_2=""
			V_SEC_DEPT_3=""
			V_SEC_TITLE_1=""
			V_SEC_TITLE_2=""
			V_SEC_TITLE_3=""
			V_PATRON_CD="01"
			V_FACULTY_C="99"
			V_NAME_T=""
			V_OFF_NAME_T=""
			V_FIRST_NM=""
			V_LAST_NM=""
			V_SALUTATION_C=""
			V_SMU_TEL=""
			V_EMP_N=""
			V_LOGON_T=""
			V_EMAIL_T=""
			ACCESS_START_DT=""
			ACCESS_END_DT=""
			'Start -- Updated by Avy - 6 Jul 09
			V_ACCESS_START_DT = ""
			V_ACCESS_END_DT = ""
			'End -- Updated by Avy - 6 Jul 09
			V_EXISTING_PATRON_CD = ""
			V_EXISTING_FACULTY_CD = ""
		END IF
	IF SubmitError<>"" THEN
			V_PRI_DEPT_C=MID(Request.form("PriDeptCd"),1,INSTR(Request.form("PriDeptCd"),"|")-1)
			V_PRI_TITLE=REQUEST.FORM("PriTitle")
			V_SEC_DEPT_1=REQUEST.FORM("SecDeptCd1")
			V_SEC_DEPT_2=REQUEST.FORM("SecDeptCd2")
			V_SEC_DEPT_3=REQUEST.FORM("SecDeptCd3")
			V_SEC_TITLE_1=REQUEST.FORM("SecTitle1")
			V_SEC_TITLE_2=REQUEST.FORM("SecTitle2")
			V_SEC_TITLE_3=REQUEST.FORM("SecTitle3")
			V_PATRON_CD=MID(Request.form("PATRONTypeCD"),1,INSTR(Request.form("PATRONTypeCD"),"|")-1)
			V_FACULTY_C=MID(Request.form("FACTypeCD"),1,INSTR(Request.form("FACTypeCD"),"|")-1)
			V_NAME_T=Request.form("DisplayNm")
			V_OFF_NAME_T=Request.form("OfficialNm")
			V_FIRST_NM=Request.form("FirstNm")
			V_LAST_NM=Request.form("LastNm")
			V_SALUTATION_C=Request.form("SaluteCd")
			V_SMU_TEL=Request.form("SMUTel")
			V_EMP_N=Request.form("EmpNum")
			V_LOGON_T=LCASE(Request.form("LOGON_T"))
			V_EMAIL_T=LCASE(Request.form("EMAIL_T"))
			
			'Start -- Updated by Avy - 6 Jul 09			
			IF Request.form("FromDte")<>"" THEN
				ACCESS_START_DT="1"				
				V_ACCESS_START_DT = Request.form("FromDte")
			ELSE
				ACCESS_START_DT=""
				V_ACCESS_START_DT =""
			END IF
			
			IF Request.form("ToDte")<>"" THEN
				ACCESS_END_DT="1"				
				V_ACCESS_END_DT = Request.form("ToDte")
			ELSE
				ACCESS_END_DT=""
				V_ACCESS_END_DT = ""
			END IF
			'End -- Updated by Avy - 6 Jul 09
	END IF
%>

<!-- #INCLUDE FILE="header.asp" -->
<SCRIPT LANGUAGE="JavaScript">

//Start -- Added by Avy - 6 Jul 09
// For calendar icon
function DateInfo()
{
   var datefield;
}

function getDateInfo(strTextBoxName)
{
   DateInfo.datefield="";
   if(showModalDialog("common/CalendarModal.asp",DateInfo,"dialogwidth:200px;dialogheight:160px;status:no;help:no")==false){
   } else
   {
     	window.document.myForm[strTextBoxName].value=DateInfo.datefield;
     	window.document.myForm[strTextBoxName].focus();
	}
}
//End -- Added by Avy - 6 Jul 09
function checkemail(stringVal){
	var str=stringVal;
	var filter=/^(\w[\w-]+(?:\.\w+)*)@((?:\w+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
	if (filter.test(str))
		testresults=true;
	else
		testresults=false;
	return (testresults)
}
function checkvalidchar(stringVal) {
	var validChars  = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-";
	var NotvalidChars  = "%*#;|";
		for( var i=0; i<stringVal.length; i++ ) { 
			if (validChars.indexOf(stringVal.charAt(i)) == -1 ) { 
				return false;
			}
		}
		return true;
}
function IsInvalidDate(ssday,ssmonth,ssyear) {
	if (ssday > 31)
		return true;
	else if ((ssmonth ==4)||(ssmonth ==6)||(ssmonth ==9)||(ssmonth ==11)) {
		if (ssday>30)
			return true;
	}
	else if (ssmonth==2) {
		if (ssday==29) {
			if ((ssyear % 4)> 0) {
				return true;
			}
		}
		else if (ssday > 28)	{
			return true;
		}
	}
	return false;
}

function D1beforeD2(d1,m1,y1,d2,m2,y2) {
	var date1=parseInt(y1*10000)+parseInt(m1*100)+parseInt(d1);
	var date2=parseInt(y2*10000)+parseInt(m2*100)+parseInt(d2);
	if (date1 < date2)
		return true;
	else 
		return false;
}

function checkdigit(charstr,charObj)
{
	var j = 0;
	var CharString="1234567890";
		if (charstr != "")
		{
			for (j=0; j<charstr.length; j++)
			{
			  if (CharString.indexOf(charstr.substring(j, j+1)) == "-1")
			  {
				return false;
			  }
			}
		}
		return true;
} 
 function checkForm(form1)
 {
	if (form1.DisplayNm.value == "")
    {
		alert("Please Input the Display Name.");
        form1.DisplayNm.focus();
        return false;
    }
	else if (form1.FirstNm.value == "")
    {
		alert("Please Input the First Name.");
        form1.FirstNm.focus();
        return false;
    }
	else if (form1.LastNm.value == "")
    {
		alert("Please Input the Last Name.");
        form1.LastNm.focus();
        return false;
    }
    else if ((form1.EmpNum.value != "")&&(form1.EmpNum.value.length<8))
	{
		alert("Employee Number should have 8 digits.");
		document.myForm.EmpNum.focus();
		return false;
	}

	if (form1.LOGON_T.value == "")
    {
		alert("Invalid Network ID format -- Please Input a valid Network ID.");
        form1.LOGON_T.focus();
        return false;
    }
	else if (!checkvalidchar(form1.LOGON_T.value)) { 
 		alert("The Network ID Field Contains Invalid Characters.");
		form1.LOGON_T.focus();
		return false;

	} 
	else if (form1.EMAIL_T.value == "")
    {
	alert("Invalid Email format -- Please Input a valid Email Address.");
        form1.EMAIL_T.focus();
        return false;
    }
	else if (!checkemail(form1.EMAIL_T.value)){ 
		alert('Please fill in a correct Email Address. e.g. michael@smu.com.sg');
		form1.EMAIL_T.focus();
		return false;
	}

	else if (form1.PatronTypeCd.value == "")
    {
		alert("Please Select the Patron Code.");
        form1.PatronTypeCd.focus();
        return false;
    }
	else if (form1.FacTypeCd.value == "")
    {
		alert("Please Select the Faculty Type.");
        form1.FacTypeCd.focus();
        return false;
    }
	else if (form1.PriDeptCd.value == "")
    {
		alert("Please Select the Primary Department.");
        form1.PriDeptCd.focus();
        return false;
    }
	else if (form1.PriTitle.value == "")
    {
		alert("Invalid Title.");
        form1.PriTitle.focus();
        return false;
    }
//Start -- Updated by Avy - 6 Jul 09
	else if (form1.FromDte.value=="") {
		alert("Invalid Start Access Date.");
		form1.FromDte.focus();
		return false;
	}	
	
	else if ((form1.FromDte.value!="")&&(isDate(form1.FromDte.value)==false))	{				
		alert("Invalid Start Access Date.");
		form1.FromDte.focus();
		return false;
    }

	else if ((form1.ToDte.value!="")&&(isDate(form1.ToDte.value)==false)) {		
		alert("Invalid End Access Date.");
		form1.ToDte.focus();
		return false;
    }	

	if ((form1.ToDte.value!="")&&(form1.FromDte.value!="")) {
		var fromDate1 = new getDate(form1.FromDte.value);
		var toDate1 = new getDate(form1.ToDte.value);
		if (fromDate1 >= toDate1) {
			alert("Invalid Period");
			form1.ToDte.focus();
			return false;
		}
	}
	
	if (form1.EMAIL_T.value!=form1.LOGON_T.value+'@smu.edu.sg') {
		if (!confirm('You have input a Non SMU formatted email address --- Continue?')) {
			form1.EMAIL_T.focus();
			return false;
		}
	}
	if ((((form1.ExistingPatronCd.value!="") && (form1.ExistingPatronCd.value!="03") && (form1.PatronTypeCd.value=="03|Y")) ||
	   ((form1.ExistingFacultyCd.value!="") && (form1.ExistingFacultyCd.value!="AJ") && (form1.FacTypeCd.value=="AJ|Y"))) &&
	   (getDate(form1.FromDte.value) > new Date())) {
		return confirm('You are about to convert full-time faculty/admin staff to adjunct/outsource staff with future start date. Please confirm to proceed.');
	}
	else {
		return confirm('Please confirm that you want save <%IF (REQUEST.QUERYSTRING("FDated")="Y") THEN %>future-dated<%ELSE%>this<%END IF%> record');
	}
}

function isDate(val) {
    var d = getDate(val);
    return !isNaN(d.valueOf());
}

function getDate(val) {
	if (val.length == 10) {
		var dteStr = val.substring(6, 10) + "-" + val.substring(3, 5) + "-" + val.substring(0,2);
		var dte = new Date(dteStr)
		return dte;
	}
	else {
		return null;
	}
}

//End -- Updated by Avy - 6 Jul 09

function setwarning(form1) {
	if ((form1.PriDeptCd.value.length>0)&&(form1.PriDeptCd.value.charAt(form1.PriDeptCd.value.length-1)=="N"))
		document.getElementById("WarningText3").innerText="* You have selected a department that that may lose the rights to maintain this record.";	
	else 
		document.getElementById("WarningText3").innerText="";

	if ((form1.FacTypeCd.value.length>0)&&(form1.FacTypeCd.value.charAt(form1.FacTypeCd.value.length-1)=="N"))
		document.getElementById("WarningText2").innerText="* You have selected a faculty type that may lose the rights to maintain this record.";	
	else 
		document.getElementById("WarningText2").innerText="";

	if ((form1.PatronTypeCd.value.length>0)&&(form1.PatronTypeCd.value.charAt(form1.PatronTypeCd.value.length-1)=="N"))
		document.getElementById("WarningText1").innerText="* You have selected a patron type that may lose the rights to maintain this record.";	
	else 
		document.getElementById("WarningText1").innerText="";

}
function adjustdepart(form1) {
	var prideptcd=form1.PriDeptCd.value.substring(0,form1.PriDeptCd.value.length-2);
	if (form1.SecDeptCd1.value==prideptcd)
		form1.SecDeptCd1.selectedIndex=0;
	if ((form1.SecDeptCd2.value==form1.SecDeptCd1.value)||(form1.SecDeptCd2.value==prideptcd))
		form1.SecDeptCd2.selectedIndex=0;
	if ((form1.SecDeptCd3.value==form1.SecDeptCd2.value)||(form1.SecDeptCd3.value==form1.SecDeptCd1.value)||(form1.SecDeptCd3.value==prideptcd))
		form1.SecDeptCd3.selectedIndex=0;
	if (form1.SecDeptCd2.selectedIndex==0) {
		form1.SecDeptCd2.selectedIndex=form1.SecDeptCd3.selectedIndex;
		form1.SecDeptCd3.selectedIndex=0;
	}
	if (form1.SecDeptCd1.selectedIndex==0) {
		form1.SecDeptCd1.selectedIndex=form1.SecDeptCd2.selectedIndex;
		form1.SecDeptCd2.selectedIndex=0;
	}
}

</script>
<br>
<div class="functionhead"><%=ValidateXSSHTML(Updatemode)%> Patron</div>
	<%IF (REQUEST.QUERYSTRING("FDated")="Y") THEN%>
		<br>
		<div style="text-align:center;">
			<span style="font-weight:bold;color:red;font-size:20px;">Future-dated Record</span>
			(click <a href="Pat_Upd.asp?ActiveNum=<%=ValidateAndEncodeXSSEx(v_ActiveNum)%>">here</a> to access the current record)
		</div>
	<%ELSE%>
		<%IF (V_HAS_FDATED_REC="Y") THEN%>
			<br>
			<div style="text-align:center;">
				<span style="font-weight:bold;color:red;font-size:20px;">This account has future-dated record</span> 
				(click <a href="Pat_Upd.asp?ActiveNum=<%=ValidateAndEncodeXSSEx(v_ActiveNum)%>&FDated=Y">here</a> to access the future-dated record)
			</div>
		<%END IF%>
	<%END IF%>
	<br/>
	<%IF request.querystring("uflg")="1" THEN%>
		<div style="padding-left:50px;color:green"><b>Record Saved</b></div>
	<%ELSEIF request.querystring("uflg")="2" THEN%>
		<div style="padding-left:50px;color:green"><b>Future-dated Record Deleted</b></div>
	<%END IF%>
	<br>
	<%IF SubmitError<>"" THEN%><div align="center" class="response_neg"><span ><b>UPDATE ERROR - <%=ValidateXSSHTML(SubmitError)%></b></span></div><br><%END IF%>
	<table border=0 align="center" bgcolor="orange" width="90%" cellpadding="3" cellspacing="1">
	<form name=myForm method=post action="pat_upd.asp?<%=ValidateXSSHTML(request.querystring)%>" >
	<tr bgcolor="orange">
		<td align="left">
			<%IF (REQUEST.QUERYSTRING("FDated")="Y") OR ((REQUEST.QUERYSTRING("FDated")<>"Y") AND (V_HAS_FDATED_REC<>"Y")) THEN%>
				<table width="100%" cellpadding="1" cellspacing="1" border="0">
					<tr>
						<td>&nbsp;
						</td>
						<td align="right">
							<b><input type="button" onClick="document.myForm.reset();" Value="Reset"  >
							<%IF (REQUEST.QUERYSTRING("FDated")="Y") THEN %>
								&nbsp; | &nbsp;
								<b><input type="submit" name="DeleteSubmit" value="Delete" onclick="return confirm('Please confirm that you want to delete the future-dated record');" >
							<%END IF%>
							&nbsp; | &nbsp;
							<b><input type="submit" name="<%=ValidateXSSHTML(UpdateMode)%>Submit" value="Save" onclick="return checkForm(document.myForm);">
						</td>
					</tr>
				</table>
			<%END IF%>
		</td>
	</tr>
	<tr><td>
		<table border=0 width="100%"  bgcolor="#FFFFFF" align="center" cellpadding="5" cellspacing="0">
<%IF V_DUMMY_DEF="I" THEN%>
			<tr bgcolor="#eddddf" > 
				<td valign="top" width="100" style="border-bottom:solid 1px #dddddd;color:darkred"><b>Account Type</td>
				<td valign="top" width="20" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" width="640" style="border-bottom:solid 1px #dddddd">
	<select name="DUMMY_F"  style="color:darkred">
		<option value="N" <%IF V_DUMMY_F="N" THEN%>selected<%END IF%>>Real Account
		<option value="Y" <%IF V_DUMMY_F="Y" THEN%>selected<%END IF%>>Dummy Account
	</select>
				</td>
			</tr>
<%ELSEIF V_DUMMY_DEF="Y" OR V_DUMMY_F="Y" THEN%>
			<tr> 
				<td valign="top" bgcolor="#eddddf" align="center" colspan="3" style="border-bottom:solid 1px #dddddd;color:darkred"><b>Dummy Account</td>
			</tr>
<%END IF%>
			<tr> 
				<td valign="top" width="100" style="border-bottom:solid 1px #dddddd"><b>Display Name</td>
				<td valign="top" width="20" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" width="640" style="border-bottom:solid 1px #dddddd">
					 <input type="text"  value="<%=ValidateXSSHTML(V_NAME_T)%>" name="DisplayNm" maxlength="50" size="50" >
				</td>
			</tr>
			<tr> 
				<td valign="top" width="100" style="border-bottom:solid 1px #dddddd"><b>Official Name</td>
				<td valign="top" width="20" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" width="640" style="border-bottom:solid 1px #dddddd">
					 <input type="text"  value="<%=ValidateXSSHTML(V_OFF_NAME_T)%>" name="OfficialNm" maxlength="200" size="100" >
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>First Name</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_FIRST_NM)%>" name="FirstNm" maxlength="80" size="50" >
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Last Name/Surname </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_LAST_NM)%>" name="LastNm" maxlength="25" size="25" >
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Salutation</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<select size="1" name="SaluteCd" >
					<%	
					'******************************
					SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
					SET CurrCmd.ACTIVECONNECTION = OBJCONN
					strSQL = "Select salutation_c,salutation_t1 from v_salutation order by salutation_t1"
					CurrCmd.CommandText = strSQL
					CurrCmd.CommandType = adCmdText
					set objRdsSC = CurrCmd.EXECUTE()


						While Not objRdsSC.EOF
					%>
						<option value="<%=ValidateXSSHTML(objRdsSC("salutation_c"))%>"  <%if objRdsSC("salutation_c") = V_SALUTATION_C then%>selected<%END IF%>><%=ValidateXSSHTML(objRdsSC("salutation_t1"))%>
					<%
						objRdsSC.MoveNext
						Wend
						Set objRdsSC = Nothing
					%>
			 		</select>
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Network ID</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					SMUSTF \ <input type="text"  value="<%=ValidateXSSHTML(V_LOGON_T)%>" name="LOGON_T" maxlength="30" size="30"  onchange="if (this.value!='') { if (!checkvalidchar(this.value)) alert('This field contains invalid characters.'); else if (this.form.EMAIL_T.value=='') this.form.EMAIL_T.value=this.value+'@smu.edu.sg';}">
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Email Address</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_EMAIL_T)%>" name="EMAIL_T" maxlength="40" size="40" >
					&nbsp;&nbsp;&nbsp;
					<input type="button"  value="&lt;&lt; SMU Email Format"  onclick="if (!checkvalidchar(this.form.LOGON_T.value)) alert('Invalid Network ID.'); else {this.form.EMAIL_T.value=this.form.LOGON_T.value+'@smu.edu.sg';}">	
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>CV/URL</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_CV_URL)%>" name="CV_URL" maxlength="500" size="100" >
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Employee Number</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_EMP_N)%>" name="EmpNum" maxlength="10" size="10" >
					<span  style="color:#777777;">&nbsp;&nbsp;* Leave Blank if Unknown</span>
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Office Tel Number</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<input type="text"  value="<%=ValidateXSSHTML(V_SMU_TEL)%>" name="SMUTel" maxlength="15" size="15" >
				</td>
			</tr>
			<input type="hidden" id="ExistingPatronCd" value="<%=ValidateXSSHTML(V_EXISTING_PATRON_CD)%>">
			<input type="hidden" id="ExistingFacultyCd" value="<%=ValidateXSSHTML(V_EXISTING_FACULTY_CD)%>">
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Patron Type</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<%

					'******************************
					SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
					SET CurrCmd.ACTIVECONNECTION = OBJCONN
					strSQL = "Select PATRON_CD,PATRON_T,'Y' AS ALLOWFLG FROM V_PATRON_CD  ORDER BY ALLOWFLG DESC,PATRON_CD"
					CurrCmd.CommandText = strSQL
					CurrCmd.CommandType = adCmdText
					set PatronCdList = CurrCmd.EXECUTE()
					%>
					<select size="1"   name="PatronTypeCd" onchange="setfaccd(this.form);setwarning(this.form);">	
					<%	While Not PatronCdList.EOF%>
						<option <%IF PatronCdList("ALLOWFLG")="N" THEN%>style="color:#777777"<%END IF%> value="<%=ValidateXSSHTML(PatronCdList("PATRON_CD")&"|"&PatronCdList("ALLOWFLG"))%>" <%if PatronCdList("PATRON_CD")=V_PATRON_CD then%>selected<%end if%>><%=ValidateXSSHTML(PatronCdList("PATRON_T"))%>
					<%
						PatronCdList.MoveNext
						Wend
					%>
      		  </select>
					<span class="response_neg"><span ID="WarningText1" >&nbsp;</span></span>
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Faculty Type</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
						<select size="1" name="FacTypeCd"  onchange="setwarning(this.form);">	
						<option value="">
     			   </select>
					<span class="response_neg"><span ID="WarningText2" >&nbsp;</span></span>
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Access Start Date</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">		
					<!-- Start -- Updated by Avy - 6 Jul 09-->
					<INPUT type=text id="FromDte" name="FromDte" size="15" maxlength="10"  value="<%=ValidateXSSHTML(V_ACCESS_START_DT)%>" >
					<!-- End -- Updated by Avy - 6 Jul 09-->
					&nbsp;&nbsp;&nbsp;&nbsp;				
					<!--Start -- Updated by Avy - 6 Jul 09-->
					<input  type="button" value="* Leave Blank if not applicable" onclick="this.form.FromDte.value='';">
					<!--End -- Updated by Avy - 6 Jul 09-->
					<br>
					&nbsp;&nbsp;&nbsp;(dd/mm/yyyy)
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #000000"><b>Access End Date</td>
				<td valign="top" style="border-bottom:solid 1px #000000"> : </td>
				<td valign="top" style="border-bottom:solid 1px #000000">
					<!-- Start -- Updated by Avy - 6 Jul 09-->
					<INPUT type=text id="ToDte" name="ToDte" size="15" maxlength="10"  value="<%=ValidateXSSHTML(V_ACCESS_END_DT)%>" >
					<!-- End -- Updated by Avy - 6 Jul 09-->
					&nbsp;&nbsp;&nbsp;&nbsp;
					<!--Start -- Updated by Avy - 6 Jul 09-->
					<input  type="button" value="* Leave Blank if not applicable" onclick="this.form.ToDte.value='';">
					<!--End -- Updated by Avy - 6 Jul 09-->
					<br>
					&nbsp;&nbsp;&nbsp;(dd/mm/yyyy)
				</td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Primary Department</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<select name="PriDeptCd"  onchange="setwarning(this.form);adjustdepart(this.form);">
					<option value="">
					<%
					'******************************
					SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
					SET CurrCmd.ACTIVECONNECTION = OBJCONN
					strSQL = "Select DEPT_C,DEPT_T, 'Y' AS ALLOWFLG from V_DEPARTMENT order by DEPT_T"
					CurrCmd.CommandText = strSQL
					CurrCmd.CommandType = adCmdText
					set DeptList = CurrCmd.EXECUTE()
					%>
					<%do while not DeptList.eof%>
						<option <%IF DeptList("ALLOWFLG")="N" THEN%>style="color:#777777"<%END IF%> value="<%=ValidateXSSHTML(DeptList("DEPT_C")&"|"&UCASE(DeptList("ALLOWFLG")))%>" <%IF V_PRI_DEPT_C=UCASE(DeptList("DEPT_C")) THEN%>selected<%END IF%>><%=ValidateXSSHTML(DeptList("DEPT_T"))%> [<%=ValidateXSSHTML(DeptList("DEPT_C"))%>]
					<%
						DeptList.movenext
						loop
					%>
					</select>
					<br>		
					&nbsp;&nbsp;&nbsp;Title: 
					<input type="text"  value="<%=ValidateXSSHTML(V_PRI_TITLE)%>" name="PriTitle" maxlength="200" size="50" >
					<div class="response_neg"><span ID="WarningText3" >&nbsp;</span></div>
				</td>
			</tr>
			<tr bgcolor="#f4f4f4"> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><b>Secondary Department(s)</td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd">
					<%
					DeptList.movefirst
					%>
					Dept:
					<select name="SecDeptCd1"  onchange="adjustdepart(this.form);">
					<option value="">
					<%do while not DeptList.eof%>
						<option value="<%=ValidateXSSHTML(DeptList("DEPT_C"))%>" <%IF V_SEC_DEPT_1=UCASE(DeptList("DEPT_C")) THEN%>selected<%END IF%>><%=ValidateXSSHTML(DeptList("DEPT_T"))%> [<%=ValidateXSSHTML(DeptList("DEPT_C"))%>]
					<%
						DeptList.movenext
						loop
					%>
					</select>
					<BR>&nbsp;&nbsp;&nbsp;Title: 
					<input type="text"  value="<%=ValidateXSSHTML(V_SEC_TITLE_1)%>" name="SecTitle1" maxlength="200" size="50" >
					<br><br>
					<%DeptList.movefirst%>
					&nbsp;&nbsp;Dept:
					<select name="SecDeptCd2"  onchange="adjustdepart(this.form);">
					<option value="">
					<%do while not DeptList.eof%>
						<option value="<%=ValidateXSSHTML(DeptList("DEPT_C"))%>" <%IF V_SEC_DEPT_2=UCASE(DeptList("DEPT_C")) THEN%>selected<%END IF%>><%=ValidateXSSHTML(DeptList("DEPT_T"))%> [<%=ValidateXSSHTML(DeptList("DEPT_C"))%>]
					<%
						DeptList.movenext
						loop
					%>
					</select>
					<BR>&nbsp;&nbsp;&nbsp;Title: 
					<input type="text"  value="<%=ValidateXSSHTML(V_SEC_TITLE_2)%>" name="SecTitle2" maxlength="200" size="50" >
					<BR>
					<br><br>
					<%DeptList.movefirst%>
					&nbsp;&nbsp;Dept:
					<select name="SecDeptCd3"  onchange="adjustdepart(this.form);">
					<option value="">
					<%do while not DeptList.eof%>
						<option value="<%=ValidateXSSHTML(DeptList("DEPT_C"))%>" <%IF V_SEC_DEPT_3=UCASE(DeptList("DEPT_C")) THEN%>selected<%END IF%>><%=ValidateXSSHTML(DeptList("DEPT_T"))%> [<%=ValidateXSSHTML(DeptList("DEPT_C"))%>]
					<%
						DeptList.movenext
						loop
					%>
					</select>
					<BR>&nbsp;&nbsp;&nbsp;Title: 
					<input type="text"  value="<%=ValidateXSSHTML(V_SEC_TITLE_3)%>" name="SecTitle3" maxlength="200" size="50" >
					<BR>
				</td>
			</tr>

			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i>Created By</i></td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i><%=ValidateXSSHTML(PatroNList("CREATION_ID"))%></i></td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i>Created On</i></td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i><%=ValidateXSSHTML(PatroNList("CREATION_DT"))%></i></td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i>Last Updated By</i></td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i><%=ValidateXSSHTML(PatroNList("LAST_UPD_ID"))%></i></td>
			</tr>
			<tr> 
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i>Last Updated On</i></td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"> : </td>
				<td valign="top" style="border-bottom:solid 1px #dddddd"><i><%=ValidateXSSHTML(PatroNList("LAST_UPD_DT"))%></i></td>
			</tr>
						
			</table>
	</td></tr>
	<tr bgcolor="orange">
		<td align="left">
			<%IF (REQUEST.QUERYSTRING("FDated")="Y") OR ((REQUEST.QUERYSTRING("FDated")<>"Y") AND (V_HAS_FDATED_REC<>"Y")) THEN%>
				<table width="100%" cellpadding="1" cellspacing="1" border="0" >
					<tr>
						<td>&nbsp;</td>
						<td align="right">
							<b><input type="button" onClick="document.myForm.reset();" Value="Reset"  >
							<%IF (REQUEST.QUERYSTRING("FDated")="Y") THEN %>
								&nbsp; | &nbsp;
								<b><input type="submit" name="DeleteSubmit" value="Delete" onclick="return confirm('Please confirm that you want to delete the future-dated record');" >
							<%END IF%>
							&nbsp; | &nbsp;
							<b><input type="submit" name="<%=ValidateXSSHTML(UpdateMode)%>Submit" value="Save" onclick="return checkForm(document.myForm);">
						</td>
					</tr>
				</table>
			<%END IF%>
		</td>
	</tr>

</form>
</table>

<script>
	function setfaccd(form1) {
		var currpattroncd=form1.PatronTypeCd.value;
		if (currpattroncd.substring(0,2)=="02") {
			<%	
			'******************************
			SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
			SET CurrCmd.ACTIVECONNECTION = OBJCONN
			strSQL = "Select FACULTY_C,FACULTY_T, 'Y' AS ALLOWFLG FROM V_FACULTY_CD where FACULTY_C NOT IN ('99') ORDER BY ALLOWFLG DESC, FACULTY_C"
			CurrCmd.CommandText = strSQL
			CurrCmd.CommandType = adCmdText
			set FacCdList = CurrCmd.EXECUTE()

				i=0
				do while not FacCdList.eof 
				%>
			<%
			i=i+1
			FacCdList.movenext 
			LOOP
			%>
			form1.FacTypeCd.options.length=<%=ValidateXSSHTML(i)%>;
			<%
				FacCdList.movefirst
				i=0
				do while not FacCdList.eof 
				%>
				<%IF FacCdList("Faculty_C")=V_FACULTY_C THEN%>
					form1.FacTypeCd.selectedIndex=<%=ValidateXSSHTML(i)%>;
				<%END IF%>
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].value="<%=ValidateXSSHTML(FacCdList("Faculty_C"))%>|<%=ValidateXSSHTML(FacCdList("ALLOWFLG"))%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].text="<%=replace(replace(ValidateXSSHTML(replace(replace(FacCdList("Faculty_T"), "<", "[LT]"), ">", "[GT]")), "[LT]", "<"), "[GT]", ">")%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].style.color="<%IF FacCdList("ALLOWFlg")="N" THEN%>#777777<%ELSE%>#000000<%END IF%>";
			<%
			i=i+1	
			FacCdList.movenext 
			LOOP
			%>
		}
		else if (currpattroncd.substring(0,2)=="03") {
			<%	
			'******************************
			SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
			SET CurrCmd.ACTIVECONNECTION = OBJCONN
			strSQL = "Select FACULTY_C,FACULTY_T, 'Y' AS ALLOWFLG FROM V_FACULTY_CD where FACULTY_C IN ('99') ORDER BY ALLOWFLG DESC, FACULTY_C"
			CurrCmd.CommandText = strSQL
			CurrCmd.CommandType = adCmdText
			set FacCdList = CurrCmd.EXECUTE()
				i=0
				do while not FacCdList.eof 
				%>
			<%
			i=i+1
			FacCdList.movenext 
			LOOP
			%>
			form1.FacTypeCd.options.length=<%=ValidateXSSHTML(i)%>;
			<%
				FacCdList.movefirst
				i=0
				do while not FacCdList.eof 
				%>
				<%IF FacCdList("Faculty_C")=V_FACULTY_C THEN%>
					form1.FacTypeCd.selectedIndex=<%=ValidateXSSHTML(i)%>;
				<%END IF%>
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].value="<%=ValidateXSSHTML(FacCdList("Faculty_C"))%>|<%=ValidateXSSHTML(FacCdList("ALLOWFLG"))%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].text="<%=replace(replace(ValidateXSSHTML(replace(replace(FacCdList("Faculty_T"), "<", "[LT]"), ">", "[GT]")), "[LT]", "<"), "[GT]", ">")%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].style.color="<%IF FacCdList("ALLOWFlg")="N" THEN%>#777777<%ELSE%>#000000<%END IF%>";
			<%
			i=i+1	
			FacCdList.movenext 
			LOOP
			%>
		}
		else {
			<%
			'******************************
			SET CurrCmd = SERVER.CREATEOBJECT("ADODB.COMMAND")
			SET CurrCmd.ACTIVECONNECTION = OBJCONN
			strSQL = "Select FACULTY_C,FACULTY_T, 'Y' AS ALLOWFLG FROM V_FACULTY_CD where FACULTY_C IN ('99') ORDER BY ALLOWFLG DESC, FACULTY_C"
			CurrCmd.CommandText = strSQL
			CurrCmd.CommandType = adCmdText
			set FacCdList = CurrCmd.EXECUTE()
				i=0
				do while not FacCdList.eof 
				%>
			<%
			i=i+1
			FacCdList.movenext 
			LOOP
			%>
			form1.FacTypeCd.options.length=<%=ValidateXSSHTML(i)%>;
			<%
				IF i>0 THEN
					FacCdList.movefirst
				END IF
				i=0
				do while not FacCdList.eof 
			%>

			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].value="<%=ValidateXSSHTML(FacCdList("Faculty_C"))%>|<%=ValidateXSSHTML(FacCdList("ALLOWFLG"))%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].text="<%=replace(replace(ValidateXSSHTML(replace(replace(FacCdList("Faculty_T"), "<", "[LT]"), ">", "[GT]")), "[LT]", "<"), "[GT]", ">")%>";
			form1.FacTypeCd.options[<%=ValidateXSSHTML(i)%>].style.color="<%IF FacCdList("ALLOWFlg")="N" THEN%>#777777<%ELSE%>#000000<%END IF%>";
				<%IF FacCdList("Faculty_C")=V_FACULTY_C THEN%>
					form1.FacTypeCd.selectedIndex=<%=ValidateXSSHTML(i)%>;
				<%END IF%>
			<%
			i=i+1
			FacCdList.movenext 
			LOOP
			%>
		}
	}

	setfaccd(document.myForm);
	setwarning(document.myForm);

</script>
<br>
<!-- #INCLUDE FILE="footer.asp" -->