<div align="center">

## LDAP AUthentication


</div>

### Description

Authenticates user using LDAP/ADSI
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Taner Akbulut](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/taner-akbulut.md)
**Level**          |Advanced
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/taner-akbulut-ldap-authentication__1-58689/archive/master.zip)





### Source Code

```
Option Explicit
Public gstrLDAPURL As String
Public Function Authenticate(strUserName As String, strPassword As String) As Boolean
  On Error Resume Next
  Dim conLDAP As ADODB.Connection
  Dim strSQL As String
  Dim strLDAPConn As String
  Dim rsUser As ADODB.Recordset
  Set conLDAP = New ADODB.Connection
  conLDAP.Provider = "ADSDSOOBject"
  strSQL = "Select AdsPath, cn From 'LDAP://" & gstrLDAPURL _
       & "' where objectClass='user'" _
       & " and objectcategory='person' and" _
       & " SamAccountName='" & strUserName & "'"
  conLDAP.Provider = "ADsDSOObject"
  conLDAP.Properties("User ID") = strUserName
  conLDAP.Properties("Password") = strPassword
  conLDAP.Properties("Encrypt Password") = True
  'open connection + password
  conLDAP.Open "DS Query", strUserName, strPassword
  'execute LDAP query
  Err.Clear
  Set rsUser = conLDAP.Execute(strSQL)
  'rs will be empty if authentication fail
  Authenticate = False
  If Err.Number = 0 Then
    If Not (rsUser Is Nothing) Then
      If Not (rsUser.EOF And rsUser.BOF) Then
        Authenticate = True
      End If
    End If
  ElseIf Err.Number = -2147217865 Then
    MsgBox "Error in LDAP settings" & vbCrLf _
        & "Call Admin"
  End If
End Function
```

