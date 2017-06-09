---
title: Document.ToggleFormsDesign Method (Word)
keywords: vbawd10.chm158007440
f1_keywords:
- vbawd10.chm158007440
ms.prod: word
api_name:
- Word.Document.ToggleFormsDesign
ms.assetid: 4db26f6c-8e59-33b6-34eb-708b39cbed9f
ms.date: 06/08/2017
---


# Document.ToggleFormsDesign Method (Word)

Switches form design mode on or off.


## Syntax

 _expression_ . **ToggleFormsDesign**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

When Microsoft Word is in form design mode, the  **Control Toolbox** toolbar is displayed. You can use this toolbar to insert Microsoft ActiveX controls such as command buttons, scroll bars, and option buttons. In form design mode, event procedures do not run, and when you click an embedded control, the control's sizing handles appear.


## Example

This example switches to form design mode if the  **Control Toolbox** toolbar is not currently displayed.


```vb
If CommandBars("Control Toolbox").Visible = False Then 
 ActiveDocument.ToggleFormsDesign 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

