---
title: View.MailMergeDataView Property (Word)
keywords: vbawd10.chm161808389
f1_keywords:
- vbawd10.chm161808389
ms.prod: word
api_name:
- Word.View.MailMergeDataView
ms.assetid: 2252ea96-70ac-f9f1-554f-59a8337c9b5c
ms.date: 06/08/2017
---


# View.MailMergeDataView Property (Word)

 **True** if mail merge data is displayed instead of mail merge fields in the specified window. Read/write **Boolean** .


## Syntax

 _expression_ . **MailMergeDataView**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

If the specified window isn't a main document, an error occurs.


## Example

If the active document includes at least one mail merge field, this example displays mail merge data from the first record in the attached data source.


```vb
If ActiveDocument.MailMerge.Fields.Count >= 1 Then 
 ActiveDocument.MailMerge.DataSource.ActiveRecord = 1 
 ActiveDocument.ActiveWindow.View.ShowFieldCodes = False 
 ActiveDocument.ActiveWindow.View.MailMergeDataView = True 
End If
```

This example switches between viewing mail merge fields and viewing the resulting data.




```vb
With ActiveDocument.ActiveWindow.View 
 .ShowFieldCodes = False 
 .MailMergeDataView = Not .MailMergeDataView 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

