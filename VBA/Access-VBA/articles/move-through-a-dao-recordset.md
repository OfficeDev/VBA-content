---
title: Move Through a DAO Recordset
ms.prod: access
ms.assetid: 7d788b60-c6e8-dea7-68fe-01b893fc3374
ms.date: 06/08/2017
---


# Move Through a DAO Recordset

A  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object usually has a current position, most often at a record. When you refer to the fields in a **Recordset**, you obtain values from the record at the current position, which is known as the current record. However, the current position can also be immediately before the first record in a **Recordset** or immediately after the last record. In certain circumstances, the current position is undefined.

You can use the following  **Move** methods to loop through the records in a **Recordset**:

- The  **[MoveFirst](http://msdn.microsoft.com/library/338F7E86-6997-B80A-FC7A-A395D10B4A62%28Office.15%29.aspx)** method moves to the first record.
    
- The  **[MoveLast](http://msdn.microsoft.com/library/FC0F7A33-1F55-9F5B-B00D-1B81F49B1C3E%28Office.15%29.aspx)** method moves to the last record.
    
- The  **[MoveNext](http://msdn.microsoft.com/library/0A1315CF-92F8-B8EF-1542-081E8C2D5BE0%28Office.15%29.aspx)** method moves to the next record.
    
- The  **[MovePrevious](http://msdn.microsoft.com/library/82A3BC3E-5221-9A1A-1350-47BC6759EDEB%28Office.15%29.aspx)** method moves to the previous record.
    
- The  **[Move](http://msdn.microsoft.com/library/21CA5AB5-FF71-1AE8-21B3-8991D5F795CF%28Office.15%29.aspx)** method moves forward or backward the number of records you specify in its syntax.
    
You can use each of these methods on table-type, dynaset-type, and snapshot-type  **Recordset** objects. On a forward-only-type **Recordset** object, you can use only the **MoveNext** and **Move** methods. If you use the **Move** method on a forward-only-type **Recordset**, the argument specifying the number of rows to move must be a positive integer.
The following code example opens a  **Recordset** object on the Employees table containing all of the records that have a **Null** value in the ReportsTo field. The function then updates the records to indicate that these employees are temporary employees. For each record in the **Recordset**, the example changes the Title and Notes fields, and saves the changes with the **[Update](http://msdn.microsoft.com/library/AAD4171A-DA95-ED72-86B3-714615EA0AC8%28Office.15%29.aspx)** method. It uses the **MoveNext** method to move to the next record.



```vb
Sub UpdateEmployees() 
 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
Dim strSQL As String 
Dim intI As Integer 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   ' Open a recordset on all records from the Employees table that have 
   ' a Null value in the ReportsTo field. 
   strSQL = "SELECT * FROM Employees WHERE ReportsTo IS NULL" 
   Set rstEmployees = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   ' If the recordset is empty, exit. 
   If rstEmployees.EOF Then Exit Sub 
 
   intI = 1 
   With rstEmployees 
      Do Until .EOF 
         .Edit 
         ![ReportsTo] = 5 
         ![Title] = "Temporary" 
         ![Notes] = rstEmployees![Notes] &; "Temp #" &; intI 
         .Update 
         .MoveNext 
         intI = intI + 1 
      Loop 
   End With 
 
   RstEmployees.Close 
   dbsNorthwind.Close 
 
   Set rstEmployees = Nothing 
   Set dbsNorthwind = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```


 **Note**  The previous example is provided only for the purposes of illustrating the  **Update** and **MoveNext** methods. For optimal performance, it is recommended that you perform this bulk operation with a SQL UPDATE query.


