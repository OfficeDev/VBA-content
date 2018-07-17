---
title: Extract Data from a Record in a DAO Recordset
ms.prod: access
ms.assetid: cd0d8c73-c9a7-3565-514d-6b379ac2d690
ms.date: 06/08/2017
---


# Extract Data from a Record in a DAO Recordset

After you have located a particular record or records, you may want to extract data to use in your application instead of modifying the underlying source table.


## Copying a Single Field

You can copy a single field of a record to a variable of the appropriate data type. The following example extracts three fields from the first record in a  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
Dim strFirstName As String 
Dim strLastName As String 
Dim strTitle As String 
 
   Set dbsNorthwind = CurrentDb 
   Set rstEmployees = dbsNorthwind.OpenRecordset("Employees") 
 
   rstEmployees.MoveFirst 
   strFirstName = rstEmployees!FirstName 
   strLastName = rstEmployees!LastName 
   strTitle = rstEmployees!Title 

```


## Copying Entire Records to an Array

To copy one or more records, you can create a two-dimensional array and copy records one at a time. You increment the first subscript for each field and the second subscript for each record.

A fast way to do this is to use the  **[GetRows](http://msdn.microsoft.com/library/59F6E4F0-E7B1-DB60-31C7-3338B66D3345%28Office.15%29.aspx)** method, which returns a two-dimensional array. The first subscript identifies the field and the second identifies the row number, as follows.




```
varRecords(intField, intRecord) 

```

The following code example uses an SQL statement to retrieve three fields from a table called Employees into a  **Recordset** object. It then uses the **GetRows** method to retrieve the first three records of the **Recordset**, and it stores the selected records in a two-dimensional array. It then prints each record, one field at a time, by using the two array indexes to select specific fields and records.

To show how the array indexes are used, the following example uses a separate statement to identify and print each field of each record. In practice, it would be more reliable to use two loops, one nested in the other, and to provide integer variables for the indexes that step through both dimensions of the array.




```vb
Sub GetRowsTest() 
 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
Dim varRecords As Variant 
Dim intNumReturned As Integer 
Dim intNumColumns As Integer 
Dim intColumn As Integer 
Dim intRow As Integer 
Dim strSQL As String 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
   strSQL = "SELECT FirstName, LastName, Title FROM Employees" 
   Set rstEmployees = dbsNorthwind.OpenRecordset(strSQL, dbOpenSnapshot) 
 
   varRecords = rstEmployees.GetRows(3) 
   intNumReturned = UBound(varRecords, 2) + 1 
   intNumColumns = UBound(varRecords, 1) + 1 
 
   For intRow = 0 To intNumReturned - 1 
      For intColumn = 0 To intNumColumns - 1 
         Debug.Print varRecords(intColumn, intRow) 
      Next intColumn 
   Next intRow 
 
   rstEmployees.Close 
   dbsNorthwind.Close 
 
   Set rstEmployees = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```

You can use subsequent calls to the  **GetRows** method if more records are available. Because the array is filled as soon as you call the **GetRows** method, you can see why this approach is much faster than copying one field at a time.

Notice also that you do not have to declare the Variant as an array, because this is done automatically when the  **GetRows** method returns records. This enables you to use fixed-length array dimensions without knowing how many records or fields will be returned, instead of using variable-length dimensions that take up more memory.

If you are trying to retrieve all the rows by using multiple  **GetRows** calls, use the **[EOF](http://msdn.microsoft.com/library/AA82C6F9-89DA-1061-437C-8FFB000744B6%28Office.15%29.aspx)** property to be sure that you are at the end of the **Recordset**. The **GetRows** method may return fewer rows than you request. If you request more than the remaining number of rows in a **Recordset**, for example, the **GetRows** method returns only the rows that remain. Similarly, if it cannot retrieve a row in the range requested, it does not return that row. For example, if the fifth record cannot be retrieved in a group of 10 records that you are trying to retrieve, the **GetRows** method returns four records and leaves the current record position on the record that caused a problem, and does not generate a run-time error. This situation may occur if a record in a dynaset was deleted by another user. If it returns fewer records than the number requested and you are not at the end of the file, you need to read each field in the current record to determine what error the **GetRows** method encountered.

Because the  **GetRows** method always returns all the fields in the **Recordset** object, you may want to create a query that returns just the fields that you need. This is especially important for OLE Object and Memo fields.


