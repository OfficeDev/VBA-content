---
title: Application.Selection Property (Publisher)
keywords: vbapb10.chm131109
f1_keywords:
- vbapb10.chm131109
ms.prod: publisher
api_name:
- Publisher.Application.Selection
ms.assetid: b4a542a7-cb54-476b-9ccf-004ce4b9ec47
ms.date: 06/08/2017
---


# Application.Selection Property (Publisher)

Returns a  **[Selection](selection-object-publisher.md)** object that represents a selected range or the cursor.


## Syntax

 _expression_. **Selection**

 _expression_A variable that represents an  **Application** object.


## Example

This example tests whether the current selection is text. If it is text, the selected text is then displayed in a message box.


```vb
Sub Selectable() 
 
 If Selection.Type = pbSelectionText Then MsgBox Selection.TextRange 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

