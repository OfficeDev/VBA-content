---
title: Window.Close Method (Word)
keywords: vbawd10.chm157417574
f1_keywords:
- vbawd10.chm157417574
ms.prod: word
api_name:
- Word.Window.Close
ms.assetid: 125fb97f-cfb0-988e-6405-56ddce68b779
ms.date: 06/08/2017
---


# Window.Close Method (Word)

Closes the specified window.


## Syntax

 _expression_ . **Close**( **_SaveChanges_** , **_RouteDocument_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|Specifies the save action for the document. Can be one of the following  **WdSaveOptions** constants: **wdDoNotSaveChanges** , **wdPromptToSaveChanges** , or **wdSaveChanges** .|
| _RouteDocument_|Optional| **Variant**| **True** to route the document to the next recipient. If the document doesn't have a routing slip attached, this argument is ignored.|

## Example

This example closes the active window of the active document and saves it.


```vb
ActiveDocument.ActiveWindow.Close SaveChanges:=wdSaveChanges
```


## See also


#### Concepts


[Window Object](window-object-word.md)

