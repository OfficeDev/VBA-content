---
title: Return a Random Record from a DAO Recordset
ms.prod: access
ms.assetid: 16d8998f-0aca-a5e6-dec4-2be93c41a595
ms.date: 06/08/2017
---


# Return a Random Record from a DAO Recordset

Access does not have a built-in mechanism for returning a random record from a set of records. This topic describes a sample user-defined function that you can use to return a random record. 


```vb
Function FindRandom(RecordSetName As String, Fieldname As String) 
 
 Dim MyDB As Database 
 Dim MyRS As Recordset 
 Dim SpecificRecord As Long, i As Long, NumOfRecords As Long 
 
 Set MyDB = CurrentDB() 
 Set MyRS = MyDB.OpenRecordset(RecordSetName, dbOpenDynaset) 
 On Error GoTo NoRecords 
 MyRS.MoveLast 
 NumOfRecords = MyRS.RecordCount 
 SpecificRecord = Int(NumOfRecords * Rnd) 
 If SpecificRecord = NumOfRecords Then 
   SpecificRecord = SpecificRecord - 1 
 End If 
 MyRS.MoveFirst 
 For i = 1 To SpecificRecord 
   MyRS.MoveNext 
 Next i 
 FindRandom = MyRS(Fieldname) 
 Exit Function 
 
NoRecords: 
 If Err = 3021 Then 
   MsgBox "There Are No Records In The Dynaset", 16, "Error" 
 Else 
   MsgBox "Error - " &; Err &; Chr$(13) &; Chr$(10) &; Error, _ 
     16, "Error" 
 End If 
 FindRandom = "No Records" 
 Exit Function 
 
End Function 
```


