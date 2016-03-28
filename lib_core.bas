' reference : Microsoft ActiveX Data Objects 2.8 library

' trouve la date d'expiration du compte AD du user username
Function AD_AccountExpirationDate(ByVal username As String) As Date
    distinguishedName = AD_UserAccDistinguishedName(username)
    On Error GoTo myerr
    Set objUser = GetObject _
    ("LDAP://" & distinguishedName)
    On Error GoTo 0
    
    dtmAccountExpiration = objUser.AccountExpirationDate
  
    If Err.Number = -2147467259 Or dtmAccountExpiration = "1/1/1970" Then
        AD_AccountExpirationDate = 0
    Else
        AD_AccountExpirationDate = dtmAccountExpiration
    End If
    Exit Function
myerr:
    Debug.Print "Can't find user " & username & "=" & distinguishedName
End Function

' trouve le nom "distinguishedName" du user username
' ça donne qqc comme : CN=BAUGE Carlo,OU=DGAF,OU=Users,OU=IDF,OU=FR,OU=Countries,DC=emea,DC=loreal,DC=intra
Function AD_UserAccDistinguishedName(ByVal username As String) As String
    Set rootDSE = GetObject("LDAP://RootDSE")
    base = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    'filter on user objects with the given account name
    fltr = "(&(objectClass=user)(objectCategory=Person)" & _
            "(sAMAccountName=" & username & "))"
    'add other attributes according to your requirements
    attr = "distinguishedname"
    scope = "subtree"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Provider = "ADsDSOObject"
    conn.Open "Active Directory Provider"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = base & ";" & fltr & ";" & attr & ";" & scope
    
    Set rs = cmd.Execute
    Do Until rs.EOF
      AD_UserAccDistinguishedName = rs.Fields("distinguishedName").Value
      rs.MoveNext
    Loop
    rs.Close
    
    conn.Close
End Function
