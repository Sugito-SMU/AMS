<%
response.buffer=true
DIM access
DIM VersionNum
access="Y"
%>

<!-- #INCLUDE FILE="../../../../../../../elmo/login_header.asp" -->
<!-- #INCLUDE FILE="../../../../../../../../elmo/prjgsacc.asp" -->    
<!-- #INCLUDE FILE="../../../../../../../../elmo/rights.asp" -->
<%
    access = "Y"

	IF (access = "N") THEN
   	RESPONSE.WRITE "Access is denied. You are not authorised to access this application."
   	RESPONSE.END
  	END IF

DeptCon=" 1=1 "
Patroncon=" 1=1 "
FacultyCon=" 1=1 "  
DummyCon=" 1=1 "
PatronSelectcon=" 1=1 "
FacultySelectCon=" 1=1 "
V_DUMMY_DEF="I"



%>
<!-- #INCLUDE FILE="../../../../../elmo/prjfams.asp" -->

<%
Sub CheckForErrors (ObjCmd,FuncName,IsInTrans)
    'This Is used when calling stored procedures To check If there are any errors.
    'All stored procedures that I design have an output Integer Called ResCSP. A non zero value implies Error.
    if CINT(Objcmd.Parameters("ResCSP").value)<>0 Then
        If IsInTrans=1 then
            objCmd.ActiveConnection.RollbackTrans
        End if
        Response.Write ("<h1>An Error has occurred</h1>: ["&Objcmd.Parameters("ResCSP").value&"] " & FuncName)
        Response.End
    End if
End Sub
%>
