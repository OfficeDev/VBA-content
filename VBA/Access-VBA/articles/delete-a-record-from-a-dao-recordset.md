---
title: Delete a Record From a DAO Recordset
ms.prod: access
ms.assetid: 7407b757-4c00-2ea7-c93f-303c09afff26
ms.date: 06/08/2017
---


# Delete a Record From a DAO Recordset

You can delete an existing record in a table or dynaset-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object by using the **[Delete](http://msdn.microsoft.com/library/310C0370-6F48-9E6B-ED1F-8D8133C52ED3%28Office.15%29.aspx)** method. You cannot delete records from a snapshot-type **Recordset** object. The following code example deletes all the duplicate records in the Shippers table.


```vb
Sub DeleteDuplicateShippers() 
 
Dim dbsNorthwind As DAO.Database 
Dim rstShippers As DAO.Recordset 
Dim strSQL As String 
Dim strName As String 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
   strSQL = "SELECT * FROM Shippers ORDER BY CompanyName, ShipperID" 
   Set rstShippers = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   ' If no records in Shippers table, exit. 
   If rstShippers.EOF Then Exit Sub 
 
   strName = rstShippers![CompanyName] 
   rstShippers.MoveNext 
 
   Do Until rstShippers.EOF 
      If rstShippers![CompanyName] = strName Then 
         rstShippers.Delete 
      Else 
         strName = rstShippers![CompanyName] 
      End If 
      rstShippers.MoveNext 
   Loop 
 
Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Function
```


When you use the  **Delete** method, the Access database engine immediately deletes the current record without any warning or prompting. Deleting a record does not automatically cause the next record to become the current record; to move to the next record you must use the **[MoveNext](http://msdn.microsoft.com/library/0A1315CF-92F8-B8EF-1542-081E8C2D5BE0%28Office.15%29.aspx)** method. Be aware that after you have moved off the deleted record, you cannot move back to it.

If you try to access a record after deleting it on a table-type  **Recordset**, you will see error 3167, "Record is deleted." On a dynaset, you will see error 3021, "No current record."
If you have a  **Recordset** clone positioned at the deleted record and you try to read its value, you will see error 3167 regardless of the type of **Recordset** object. Trying to use a bookmark to move to a deleted record will also result in error 3167.

