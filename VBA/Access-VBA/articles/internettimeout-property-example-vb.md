---
title: InternetTimeout Property Example (VB)
ms.prod: access
ms.assetid: 095a384d-5c02-a096-d8f8-31edbc941f90
ms.date: 06/08/2017
---


# InternetTimeout Property Example (VB)

  

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [InternetTimeout](http://msdn.microsoft.com/library/66fc6e87-3d23-ce2c-18f5-0fc83ac43801%28Office.15%29.aspx) property, which exists on the[DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) and[DataSpace](http://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) objects. This example uses the **DataControl** object and sets the timeout to 20 seconds.




```VB.net

'BeginInternetTimeoutVB 
 
Public Sub Main()On Error GoTo ErrorHandler 
Dim dc As RDS.DataControlDim rst As ADODB.Recordset
Set dc = New RDS.DataControl 
dc.Server = "http://MyServer"dc.ExecuteOptions = 1
dc.FetchOptions = 1dc.Connect = "Provider='sqloledb';Data Source='MySqlServer';" &; _
"Initial Catalog='Pubs';Integrated Security='SSPI';"dc.SQL = "SELECT * FROM Authors"
' Wait at least 20 secondsdc.InternetTimeout = 200 
dc.Refresh' Use another Recordset as a convenience
Set rst = dc.RecordsetDo While Not rst.EOF
Debug.Print rst!au_fname &; " " &; rst!au_lnamerst.MoveNext
Loop 
If rst.State = adStateOpen Then rst.CloseSet rst = Nothing
Set dc = NothingExit Sub 
ErrorHandler:' clean up
If Not rst Is Nothing ThenIf rst.State = adStateOpen Then rst.Close
End IfSet rst = Nothing
Set dc = Nothing 
If Err <> 0 ThenMsgBox Err.Source &; "-->" &; Err.Description, , "Error"
End If 
End Sub'EndInternetTimeoutVB
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

