<%
response.buffer=true
DIM access
DIM VersionNum
access="Y"
%>

<!-- #INCLUDE FILE="sso.inc"  -->

<%
if (Session("ADFS_USERNAME") = "") then
    if (GetSSOToken() <> "") then
        Response.Redirect("sso.asp")
    else
        Response.Redirect("sso/login.aspx?route=SSO")
    end if
else
    username = Session("ADFS_USERNAME")
end if

Response.Write username
%>
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
