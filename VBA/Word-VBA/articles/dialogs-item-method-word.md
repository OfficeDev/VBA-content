---
title: Dialogs.Item Method (Word)
keywords: vbawd10.chm152043520
f1_keywords:
- vbawd10.chm152043520
ms.prod: word
api_name:
- Word.Dialogs.Item
ms.assetid: 8a7826ce-a5b9-e0af-29cb-5dea299ab266
ms.date: 06/08/2017
---


# Dialogs.Item Method (Word)

Returns a dialog in Microsoft Word.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[Dialogs](dialogs-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdWordDialog**|A constant that specifies the dialog.|

### Return Value

Dialog


## Example

This example displays the Page Setup dialog.


```vb
Sub DialogItem() 
 Application.Dialogs.Item(wdDialogFileDocumentLayout).Display 
End Sub
```


## See also


#### Concepts


[Dialogs Collection Object](dialogs-object-word.md)

