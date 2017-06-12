---
title: TextFrame.ValidLinkTarget Method (Word)
keywords: vbawd10.chm162665358
f1_keywords:
- vbawd10.chm162665358
ms.prod: word
api_name:
- Word.TextFrame.ValidLinkTarget
ms.assetid: 09e900c9-30d8-0098-6ad1-d8c4fbaeb3cf
ms.date: 06/08/2017
---


# TextFrame.ValidLinkTarget Method (Word)

Determines whether the text frame of one shape can be linked to the text frame of another shape. .


## Syntax

 _expression_ . **ValidLinkTarget**( **_TargetTextFrame_** )

 _expression_ Required. A variable that represents a **[TextFrame](textframe-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TargetTextFrame_|Required| **TextFrame**|The target text frame to which you want to link the text frame returned by expression.|

### Return Value

Boolean


## Remarks

This method returns  **True** if TargetTextFrame is a valid target and returns **False** if TargetTextFrame already contains text, is already linked, or if the shape doesn't support attached text.


## Example

This example checks to see whether the text frames for the first and second shapes in the active document can be linked to one another. If so, the example links the two text frames.


```vb
Dim textFrame1 As TextFrame 
Dim textFrame2 As TextFrame 
 
Set textFrame1 = ActiveDocument.Shapes(1).TextFrame 
Set textFrame2 = ActiveDocument.Shapes(2).TextFrame 
If textFrame1.ValidLinkTarget(textFrame2) = True Then 
 textFrame1.Next = textFrame2 
End If
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

