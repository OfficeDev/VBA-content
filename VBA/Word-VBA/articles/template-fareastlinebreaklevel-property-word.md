---
title: Template.FarEastLineBreakLevel Property (Word)
keywords: vbawd10.chm157941774
f1_keywords:
- vbawd10.chm157941774
ms.prod: word
api_name:
- Word.Template.FarEastLineBreakLevel
ms.assetid: 4bf3fd26-0b6c-f970-19bf-1bd9d8441d54
ms.date: 06/08/2017
---


# Template.FarEastLineBreakLevel Property (Word)

Returns or sets the line break control level for the specified document. Read/write  **WdFarEastLineBreakLevel** .


## Syntax

 _expression_ . **FarEastLineBreakLevel**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


## Remarks

This property is ignored if the  **FarEastLineBreakControl** property is set to **False** .


## Example

This example sets Microsoft Word to perform line breaking on first-level kinsoku characters in the active document.


```vb
ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana
```


## See also


#### Concepts


[Template Object](template-object-word.md)

