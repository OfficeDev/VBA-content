---
title: Handler Property Example (VB)
ms.prod: access
ms.assetid: e401e7b2-754b-a66c-bfcc-8f6e3966a908
ms.date: 06/08/2017
---


# Handler Property Example (VB)

  

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [RDS DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object[Handler](http://msdn.microsoft.com/library/aaf8c8c6-f95b-3cf3-b3f6-203f37464c87%28Office.15%29.aspx) property. (See[DataFactory Customization](http://msdn.microsoft.com/library/43cd7416-1f05-87ee-22f0-6cf0d2d1b39f%28Office.15%29.aspx) for more details.)

Assume that the following sections in the parameter file, Msdfmap.ini, are located on the server:



```sql
 
[connect AuthorDataBase] 
Access=ReadWrite 
Connect="DSN=Pubs" 
[sql AuthorById] 
SQL="SELECT * FROM Authors WHERE au_id = ?" 

```

Your code looks like the following. The command assigned to the [SQL](sql-property-ado.md) property will match the ** _AuthorById_** identifier and will retrieve a row for author Michael O'Leary. The **DataControl** object **Recordset** property is assigned to a disconnected[Recordset](http://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object purely as a coding convenience.



```VB.net

'BeginHandlerVBPublic Sub Main()
On Error GoTo ErrorHandlerDim dc As New DataControl
Dim rst As ADODB.Recordsetdc.Handler = "MSDFMAP.Handler"
dc.ExecuteOptions = 1dc.FetchOptions = 1
dc.Server = "http://MyServer"dc.Connect = "Data Source=AuthorDataBase"
dc.SQL = "AuthorById('267-41-2394')"dc.Refresh 'Retrieve the record
Set rst = dc.Recordset 'Use another Recordset as a convenienceDebug.Print "Author is '" &; rst!au_fname &; " " &; rst!au_lname &; "'"
' clean upIf rst.State = adStateOpen Then rst.Close
Set rst = NothingSet dc = Nothing
Exit SubErrorHandler:
' clean upIf Not rst Is Nothing Then
If rst.State = adStateOpen Then rst.CloseEnd If
Set rst = NothingSet dc = Nothing
If Err <> 0 ThenMsgBox Err.Source &; "-->" &; Err.Description, , "Error"
End IfEnd Sub
'EndHandlerVB
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

