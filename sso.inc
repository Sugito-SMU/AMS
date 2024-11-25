<%
    
    ' This function is workaround to get SSO_TOKEN cookie
    ' ASP for unknown reason threw error 80004005 if there is a malformed cookie (e.g. without key)
    ' within a collection of cookies

    Function GetSSOToken()
        cookies = Request.ServerVariables("HTTP_COOKIE")
        token = ""
        set tokenRegex = New RegExp
        tokenRegex.Pattern = "SSO_TOKEN=((\w\w-)+\w\w)"
        tokenRegex.Global = True
        if tokenRegex.Test(cookies) Then
            set matches = tokenRegex.Execute(cookies)
            token = matches.Item(0).Submatches(0)
        end if
        GetSSOToken = token
    End Function

    Function ValidateAndRouteSSO(defaultPage, ssoPageRoute)

        set threadUtil = Server.CreateObject("ASPUtilities.ThreadUtil")
        res = threadUtil.RevertToProcessUser()

        dim mode
        dim token
        dim tokenValue
        dim certThumbprint
        dim symmKey
        ' -- mode is either "AES" or "X509"
        mode = "AES"

        ' -- certThumbprint is used if mode is X509	
        certThumbprint = "D1EB41E4DD6175724A056A11FB0A61077731E24E"

        ' -- symmKey must be 16 chars, and used if mode is AES
        symmKey = "1234567890123456"

        token = GetSSOToken()

        if (token <> "") then
            if (mode = "X509") then
                set util = Server.CreateObject("SSOUtilities.X509TokenInterop")
                tokenValue = util.ValidateAndGetToken("AMS", token, certThumbprint)
            else
                if (mode = "AES") then
                    set util = Server.CreateObject("SSOUtilities.AESTokenInterop")
                    tokenValue = util.ValidateAndGetToken("AMS", token, symmKey)
                end if
            end if
            if (tokenValue <> "") then

                Session("ADFS_USERNAME") = tokenValue
                Response.Redirect(defaultPage)
            else
                Response.Redirect(ssoPageRoute)
            end if
        else
            Response.Write "The system has encountered some error. We apologize for any inconvenience caused.<br />" &_
                        "Please contact Technology Help Centre at 6828 0123 (helpdesk@smu.edu.sg)."
            Response.End
        end if
    End Function
    
%>