---
title: Find a Record in a Dynaset-Type or Snapshot-Type DAO Recordset
ms.prod: access
ms.assetid: f79f47e1-63a9-774d-4d07-32759ac30c8b
ms.date: 06/08/2017
---


# Find a Record in a Dynaset-Type or Snapshot-Type DAO Recordset

You can use the Find methods to locate a record in a dynaset-type or snapshot-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object. DAO provides the following Find methods:


- The  **[FindFirst](http://msdn.microsoft.com/library/5FCF78CD-7D2C-2E47-14E5-996F2E14FF51%28Office.15%29.aspx)** method finds the first record that satisfies the specified criteria.
    
- The  **[FindLast](http://msdn.microsoft.com/library/65236519-3474-A760-99BC-2E8F6BFEEE7A%28Office.15%29.aspx)** method finds the last record that satisfies the specified criteria.
    
- The  **[FindNext](http://msdn.microsoft.com/library/5457DFC8-E561-5624-74D0-34278BA2E7CB%28Office.15%29.aspx)** method finds the next record that satisfies the specified criteria.
    
- The  **[FindPrevious](http://msdn.microsoft.com/library/62F26B0B-F3F1-A6FE-E84D-F93623E1F7F9%28Office.15%29.aspx)** method finds the previous record that satisfies the specified criteria.
    

When you use the Find methods, you specify the search criteria, which is typically an expression that equates a field name with a specific value.

You can locate the matching records in reverse order by finding the last occurrence with the  **FindLast** method and then using the **FindPrevious** method instead of the **FindNext** method.
DAO sets the  **[NoMatch](http://msdn.microsoft.com/library/47D03575-F570-89B5-A20F-A3BD8B8B5C6D%28Office.15%29.aspx)** property to **True** when a Find method fails and the current record position is undefined. There may be a current record, but there is no way to tell which one. To return to the previous current record following a failed Find method, use a bookmark.
The  **NoMatch** property is **False** when the operation succeeds. In this case, the current record position is the record found by one of the Find methods.
The following example illustrates how you can use the  **FindNext** method to find all orders in the Orders table that have no corresponding records in the Order Details table. The function searches for missing orders and, if it finds one, it adds the value in the OrderID field to the array aryOrders().



```vb
Function FindOrders() As Variant 
 
Dim dbsNorthwind As DAO.Database 
Dim rstOrders As DAO.Recordset 
Dim rstOrderDetails As DAO.Recordset 
Dim strSQL As String 
Dim intIndex As Integer 
Dim aryOrders() As Long 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   ' Open recordsets on the Orders and OrderDetails tables. If there are 
   ' no records in either table, exit the function. 
   strSQL = "SELECT * FROM Orders ORDER BY OrderID" 
   Set rstOrders = dbsNorthwind.OpenRecordset(strSQL, dbOpenSnapshot) 
   If rstOrders.EOF Then Exit Function 
 
   strSQL = "SELECT * FROM [Order Details] ORDER BY OrderID" 
   Set rstOrderDetails = dbsNorthwind.OpenRecordset(strSQL, _ 
                         dbOpenSnapshot) 
 
   ' For the first record in Orders, find the first matching record 
   ' in OrderDetails. If no match, redimension the array of order IDs and 
   ' add the order ID to the array. 
   intIndex = 1 
   rstOrderDetails.FindFirst "OrderID = " &; rstOrders![OrderID] 
   If rstOrderDetails.NoMatch Then 
      ReDim Preserve aryOrders (1 To intIndex) 
      aryOrders (intIndex) = rstOrders![OrderID] 
      rstOrders.MoveNext 
   End If 
 
   ' The first match has already been found, so use the FindNext method to 
   ' find the next record that satisfies the criteria. 
   Do Until rstOrders.EOF 
      rstOrderDetails.FindNext "OrderID = " &; rstOrders![OrderID] 
      If rstOrderDetails.NoMatch Then 
         intIndex = intIndex + 1 
         ReDim Preserve aryOrders (1 To intIndex) 
         aryOrders (intIndex) = rstOrders![OrderID] 
      End If 
      rstOrders.MoveNext 
   Loop 
 
   FindOrders = aryOrders 
 
   rstOrders.Close 
   rstOrderDetails.Close 
   dbsNorthwind.Close 
 
   Set rstOrders = Nothing 
   Set rstOrderDetails = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Function 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Function
```


