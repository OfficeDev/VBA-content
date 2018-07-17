---
title: Document.FormsDesign Property (Word)
keywords: vbawd10.chm158007396
f1_keywords:
- vbawd10.chm158007396
ms.prod: word
api_name:
- Word.Document.FormsDesign
ms.assetid: f5ec4968-fb3e-5cca-de0b-55c36a7ae584
ms.date: 06/08/2017
---


# Document.FormsDesign Property (Word)

 **True** if the specified document is in form design mode. Read-only **Boolean** .


## Syntax

 _expression_ . **FormsDesign**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **FormsDesign** property always returns **False** if it's used in code that is run from Microsoft Word, but it returns the correct value if it is run through Automation. For example, if you run the example from Microsoft Excel, it will return **True** if you're in design mode.

When Word is in form design mode, the  **Control Toolbox** toolbar is displayed. You can use this toolbar to insert ActiveX controls such as command buttons, scroll bars, and option buttons. In form design mode, event procedures don't run, and when you click an embedded control, the control's sizing handles appear.


## Example

This example displays a message box that indicates whether the active document is in form design mode. This example will always return  **False** .


```
Msgbox ActiveDocument.FormsDesign
```


## See also


#### Concepts


[Document Object](document-object-word.md)

