---
title: MailMergeDataSource.FieldNames Property (Word)
keywords: vbawd10.chm152895498
f1_keywords:
- vbawd10.chm152895498
ms.prod: word
api_name:
- Word.MailMergeDataSource.FieldNames
ms.assetid: 3e88ee90-c44e-1dbb-dcfd-6ea99cbb1c2c
ms.date: 06/08/2017
---


# MailMergeDataSource.FieldNames Property (Word)

Returns a  **[MailMergeFieldNames](mailmergefieldnames-object-word.md)** collection that represents the names of all the fields in the specified mail merge data source. Read-only.


## Syntax

 _expression_ . **FieldNames**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of the first field in the data source attached to the active mail merge main document.


```vb
MsgBox ActiveDocument.MailMerge.DataSource.FieldNames(1).Name
```

This example uses the mNames() array to store the names of each merge field contained in the data source attached to the active document.




```vb
Dim mNames As Variant 
Dim mmTemp As MailMerge 
Dim intCount As Integer 
Dim intIncrement As Integer 
Dim mmfnLoop As MailMergeFieldName 
 
Set mmTemp = ActiveDocument.MailMerge 
intCount = _ 
 ActiveDocument.MailMerge.DataSource.FieldNames.Count - 1 
 
ReDim mNames(intCount) 
intIncrement = 0 
 
For Each mmfnLoop In mmTemp.DataSource.FieldNames 
 mNames(intIncrement) = mmfnLoop.Name 
 intIncrement = intIncrement + 1 
Next mmfnLoop
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

