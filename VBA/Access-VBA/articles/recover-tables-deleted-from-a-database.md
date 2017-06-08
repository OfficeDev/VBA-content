---
title: Recover Tables Deleted from a Database
ms.prod: access
ms.assetid: 4d370adb-741f-269d-8def-bccec1f335f1
ms.date: 06/08/2017
---


# Recover Tables Deleted from a Database

This topic shows how to create a sample Visual Basic for Applications (VBA) function that you can use to recover tables deleted from an Access database under the following conditions: 


- The database has not been closed since the tables were deleted.
    
- The database has not been compacted since the tables were deleted.
    
- The tables were deleted using the Access user interface.
    
- The table does not contain any multivalue or Attachment fields.
    

Paste the following procedure into a standard module. 




```vb
Sub RecoverDeletedTable() 
On Error GoTo ExitHere 
 
 Dim db As DAO.Database 
 Dim strTableName As String 
 Dim strSQL As String 
 Dim intCount As Integer 
 Dim blnRestored As Boolean 
 
 Set db = CurrentDb() 
 
 For intCount = 0 To db.TableDefs.Count - 1 
 strTableName = db.TableDefs(intCount).Name 
 If Left(strTableName, 4) = "~tmp" Then 
 strSQL = "SELECT DISTINCTROW [" &; strTableName &; "].* INTO " &; Mid(strTableName, 5) &; " FROM [" &; strTableName &; "];" 
 DoCmd.SetWarnings False 
 DoCmd.RunSQL strSQL 
 MsgBox "A deleted table has been restored, using the name '" &; Mid(strTableName, 5) &; "'", vbOKOnly, "Restored" 
 blnRestored = True 
 End If 
 Next intCount 
 
 If blnRestored = False Then 
MsgBox "No recoverable tables found", vbOKOnly 
 End If 
 
'*EXIT/ERROR* 
ExitHere: 
 DoCmd.SetWarnings True 
 Set db = Nothing 
 Exit Sub 
 
ErrorHandler: 
 MsgBox Err.Description 
 Resume ExitHere 
 
End Sub
```


