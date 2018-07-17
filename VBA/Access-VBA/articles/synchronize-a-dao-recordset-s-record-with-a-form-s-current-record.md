---
title: Synchronize a DAO Recordset's Record with a Form's Current Record
ms.prod: access
ms.assetid: 2960dd7d-4c60-4148-ef58-dd44f1042851
ms.date: 06/08/2017
---


# Synchronize a DAO Recordset's Record with a Form's Current Record

The following code example uses the  **[RecordsetClone](form-recordsetclone-property-access.md)** property and the **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object to synchronize a recordset's record with the form's current record. When a company name is selected from a combo box, the **[FindFirst](http://msdn.microsoft.com/library/5FCF78CD-7D2C-2E47-14E5-996F2E14FF51%28Office.15%29.aspx)** method is used to locate the record for that company and the **Recordset** object's **[Bookmark](http://msdn.microsoft.com/library/C4B1C2D9-668E-E365-544C-EFB4AE4EFCC9%28Office.15%29.aspx)** property is assigned to the form's **[Bookmark](form-bookmark-property-access.md)** property, causing the form to display the found record.


```vb
Sub SupplierID_AfterUpdate() 
    Dim rst As Recordset 
    Dim strSearchName As String 
 
    Set rst = Me.RecordsetClone 
    strSearchName = Str(Me!SupplierID) 
    rst.FindFirst "SupplierID = " &; strSearchName 
        If rst.NoMatch Then 
            MsgBox "Record not found" 
        Else 
            Me.Bookmark = rst.Bookmark 
        End If 
    rst.Close 
End Sub
```


