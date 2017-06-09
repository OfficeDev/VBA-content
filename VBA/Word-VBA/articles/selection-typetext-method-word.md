---
title: Selection.TypeText Method (Word)
keywords: vbawd10.chm158663163
f1_keywords:
- vbawd10.chm158663163
ms.prod: word
api_name:
- Word.Selection.TypeText
ms.assetid: fb8e58cc-0c49-0efa-d60a-8be6c3d4435c
ms.date: 06/08/2017
---


# Selection.TypeText Method (Word)

Inserts the specified text.


## Syntax

 _expression_ . **TypeText**( **_Text_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

If the  **ReplaceSelection** property is **True** , the selection is replaced by the specified text. If **ReplaceSelection** is **False** , the specified text is inserted before the selection.


## Example

If Word is set so that typing replaces selected text, this example collapses the selection before inserting "Hello." This technique prevents existing document text from being replaced.


```vb
If Options.ReplaceSelection = True Then 
 Selection.Collapse Direction:=wdCollapseStart 
 Selection.TypeText Text:="Hello" 
End If
```

This example inserts "Title" followed by a new paragraph.




```vb
Options.ReplaceSelection = False 
With Selection 
 .TypeText Text:="Title" 
 .TypeParagraph 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

