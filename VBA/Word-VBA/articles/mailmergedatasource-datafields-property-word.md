---
title: MailMergeDataSource.DataFields Property (Word)
keywords: vbawd10.chm152895499
f1_keywords:
- vbawd10.chm152895499
ms.prod: word
api_name:
- Word.MailMergeDataSource.DataFields
ms.assetid: 613c4bc6-bd87-fbdc-2170-8a1daf2cfd2c
ms.date: 06/08/2017
---


# MailMergeDataSource.DataFields Property (Word)

Returns a  **[MailMergeDataFields](mailmergedatafields-object-word.md)** collection that represents the fields in the specified mail merge data source. Read-only.


## Syntax

 _expression_ . **DataFields**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of each field in the data source attached to the active mail merge main document.


```vb
Dim mmdfTemp As MailMergeDataField 
 
For Each mmdfTemp In _ 
 ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox mmdfTemp.Name 
Next mmdfTemp
```

This example displays the value of the LastName field from the first record in the data source attached to "Main.doc."




```vb
With Documents("Main.doc").MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 MsgBox .DataFields("LastName").Value 
End With
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

