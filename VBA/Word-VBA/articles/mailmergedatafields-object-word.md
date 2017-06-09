---
title: MailMergeDataFields Object (Word)
ms.prod: word
ms.assetid: a660288d-1a2c-53ec-20d2-c52353be90c8
ms.date: 06/08/2017
---


# MailMergeDataFields Object (Word)

A collection of  **[MailMergeDataField](mailmergedatafield-object-word.md)** objects that represent the data fields in a mail merge data source.


## Remarks

Use the  **DataFields** property to return the **MailMergeDataFields** collection. The following example displays the names of all the fields in the attached data source.


```vb
For Each afield In ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox afield.Name 
Next afield
```

You cannot add fields to the  **MailMergeDataFields** collection. When a data field is added to a data source, the field is automatically included in the **MailMergeDataFields** collection. Use the **EditDataSource** method to edit the contents of a data source. The following example adds a data field named "Author" to a table in the attached data source.




```vb
If ActiveDocument.MailMerge.DataSource.Type = _ 
 wdMergeInfoFromWord Then 
 ActiveDocument.MailMerge.EditDataSource 
 With ActiveDocument.Tables(1) 
 .Columns.Add 
 .Cell(Row:=1, Column:=.Columns.Count).Range.Text = "Author" 
 End With 
End If
```

Use  **DataFields** (Index), where Index is the data field name or the index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. The following example retrieves the first value from the FName field in the data source attached to the active document.




```
first = ActiveDocument.MailMerge _ 
 .DataSource.DataFields("FName").Value
```

The following example displays the name of first data field in the data source attached to the active document.




```vb
MsgBox ActiveDocument.MailMerge.DataSource.DataFields(1).Name
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

