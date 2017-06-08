---
title: Add "(All)" to a Combo Box or List Box
ms.prod: access
ms.assetid: f246db25-84b1-736f-8a79-16b9eea9cbda
ms.date: 06/08/2017
---


# Add "(All)" to a Combo Box or List Box

When you use a list box or combo box to enter selection criteria, you might want to be able to specify all records. The AddAllToList procedure illustrates how to add an  **(All)** entry at the top of a combo box.

To utilize the AddAllToList procedure, you must set the  **RowSourceType** property of the combo box or list box to **AddAllToList**.

You can specify a different item than  **(All)** to add to the list by setting the **Tag** property of the combo box or list box. For example, you can add **<None>** to the top of the list by setting the value of the **Tag** property to `1;<None>`.




```vb
Function AddAllToList(ctl As Control, lngID As Long, lngRow As Long, _ 
lngCol As Long, intCode As Integer) As Variant 
 
Static dbs As Database, rst As Recordset 
Static lngDisplayID As Long 
Static intDisplayCol As Integer 
Static strDisplayText As String 
Dim intSemiColon As Integer 
 
On Error GoTo Err_AddAllToList 
Select Case intCode 
Case acLBInitialize 
' See if function is already in use. 
If lngDisplayID <> 0 Then 
MsgBox "AddAllToList is already in use by another control!" 
AddAllToList = False 
 
Exit Function 
End If 
 
' Parse the display column and display text from Tag property. 
intDisplayCol = 1 
strDisplayText = "(All)" 
If ctl.Tag <> "" Then 
intSemiColon = InStr(ctl.Tag, ";") 
If intSemiColon = 0 Then 
intDisplayCol = Val(ctl.Tag) 
Else 
intDisplayCol = Val(Left(ctl.Tag, intSemiColon - 1)) 
strDisplayText = Mid(ctl.Tag, intSemiColon + 1) 
 
End If 
End If 
 
' Open the recordset defined in the RowSource property. 
Set dbs = CurrentDb 
Set rst = dbs.OpenRecordset(ctl.RowSource, dbOpenSnapshot) 
 
' Record and return the lngID for this function. 
lngDisplayID = Timer 
AddAllToList = lngDisplayID 
 
Case acLBOpen 
AddAllToList = lngDisplayID 
 
Case acLBGetRowCount 
' Return number of rows in recordset. 
On Error Resume Next 
 
rst.MoveLast 
AddAllToList = rst.RecordCount + 1 
 
Case acLBGetColumnCount 
' Return number of fields (columns) in recordset. 
AddAllToList = rst.Fields.Count 
 
Case acLBGetColumnWidth 
AddAllToList = -1 
 
Case acLBGetValue 
If lngRow = 0 Then 
If lngCol = intDisplayCol - 1 Then 
AddAllToList = strDisplayText 
Else 
AddAllToList = Null 
End If 
Else 
 
rst.MoveFirst 
rst.Move lngRow - 1 
AddAllToList = rst(lngCol) 
End If 
Case acLBEnd 
lngDisplayID = 0 
rst.Close 
End Select 
 
Bye_AddAllToList: 
Exit Function 
 
Err_AddAllToList: 
MsgBox Err.Description, vbOKOnly + vbCritical, "AddAllToList" 
AddAllToList = False 
Resume Bye_AddAllToList 
End Function
```


