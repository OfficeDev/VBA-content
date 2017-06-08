---
title: Add a Record to a DAO Recordset
ms.prod: access
ms.assetid: b6366906-4b37-0d35-cfd5-d38e7717131c
ms.date: 06/08/2017
---


# Add a Record to a DAO Recordset

You can add a new record to a table-type or dynaset-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object by using the **[AddNew](http://msdn.microsoft.com/library/18CB35F6-8652-FB20-2460-3D13FAE39D23%28Office.15%29.aspx)** method.


### To add a record to a table-type or dynaset-type Recordset object:


1. Use the  **AddNew** method to create a record you can edit.
    
2. Assign values to each of the record's fields.
    
3. Use the  **[Update](http://msdn.microsoft.com/library/AAD4171A-DA95-ED72-86B3-714615EA0AC8%28Office.15%29.aspx)** method to save the new record.
    
The following code example adds a record to a table-type  **Recordset** called Shippers.



```vb
Dim dbsNorthwind As DAO.Database 
Dim rstShippers As DAO.Recordset 
 
   Set dbsNorthwind = CurrentDb 
   Set rstShippers = dbsNorthwind.OpenRecordset("Shippers") 
 
   rstShippers.AddNew 
   rstShippers!CompanyName = "Global Parcel Service" 
      . 
      . ' Set remaining fields. 
      . 
 
   rstShippers.Update 

```

When you use the  **AddNew** method, the Access database engine prepares a new, blank record and makes it the current record. When you use the **Update** method to save the new record, the record that was current before you used the **AddNew** method becomes the current record again.
The new record's position in the  **Recordset** depends on whether you added the record to a dynaset-type or a table-type **Recordset** object. If you add a record to a dynaset-type **Recordset**, the new record appears at the end of the **Recordset**, no matter how the **Recordset** is sorted. To force the new record to appear in its properly sorted position, you can either use the **[Requery](http://msdn.microsoft.com/library/A5D66EB5-499C-4133-F6C3-C7A1619A8A11%28Office.15%29.aspx)** method or recreate the **Recordset** object.
If you add a record to a table-type Recordset, the record appears positioned according to the current index, or at the end of the table if there is no current index. Because the Access databse engine allows multiple users to create records in a table simultaneously, your record may not appear at the end of the  **Recordset**. Be sure to use the **[LastModified](http://msdn.microsoft.com/library/7386F25B-BDE1-A446-E980-640696A3BFEC%28Office.15%29.aspx)** property rather than the **[MoveLast](http://msdn.microsoft.com/library/FC0F7A33-1F55-9F5B-B00D-1B81F49B1C3E%28Office.15%29.aspx)** method to move to the record you just added.

 **Note**   If you use the **AddNew** method to add a record, and then move to another record or close the **Recordset** object without first using the **Update** method, your changes are lost without warning. For example, omitting the **Update** method from the preceding example results in no changes being made to the Shippers table.


