---
title: Selection.InsertBefore Method (Word)
keywords: vbawd10.chm158662758
f1_keywords:
- vbawd10.chm158662758
ms.prod: word
api_name:
- Word.Selection.InsertBefore
ms.assetid: 05dfc75f-9bb3-e090-9b31-aeb48b6c2ed8
ms.date: 06/08/2017
---


# Selection.InsertBefore Method (Word)

Inserts the specified text before the specified selection. .


## Syntax

 _expression_ . **InsertBefore**( **_Text_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

After using this method to insert text, the selection is expanded to include the new text. If the selection is a bookmark, the bookmark is also expanded to include the next text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertBefore** method. You can also use the following Visual Basic constants: **vbCr** , **vbLf** , **vbCrLf** and **vbTab** .


## Example

This example inserts the text "Hamlet" (enclosed in quotation marks) before the selection and then collapses the selection.


```vb
With Selection 
 .InsertBefore Chr(34) &; "Hamlet" &; Chr(34) &; Chr(32) 
 .Collapse Direction:=wdCollapseEnd 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

