---
title: Document.Redo Method (Word)
keywords: vbawd10.chm158007413
f1_keywords:
- vbawd10.chm158007413
ms.prod: word
api_name:
- Word.Document.Redo
ms.assetid: 0fb5671e-c933-50e6-e1fa-fe146666ad80
ms.date: 06/08/2017
---


# Document.Redo Method (Word)

Redoes the last action that was undone (reverses the  **Undo** method). Returns **True** if the actions were redone successfully.


## Syntax

 _expression_ . **Redo**( **_Times_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Times_|Optional| **Variant**|The number of actions to be redone.|

### Return Value

Boolean


## Example

This example redoes the last two actions in the Sales.doc redo list.


```
Documents("Sales.doc").Redo 2
```

This example redoes the last action in the active document. If the action is successfully redone, a message is displayed in the status bar.




```vb
On Error Resume Next 
If ActiveDocument.Redo = False Then _ 
 StatusBar = "Redo was unsuccessful"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

