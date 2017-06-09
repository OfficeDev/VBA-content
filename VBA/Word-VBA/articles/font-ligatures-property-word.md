---
title: Font.Ligatures Property (Word)
keywords: vbawd10.chm156369070
f1_keywords:
- vbawd10.chm156369070
ms.prod: word
api_name:
- Word.Font.Ligatures
ms.assetid: f1b0ff39-5eb5-e5a3-e0ff-3e88639670f9
ms.date: 06/08/2017
---


# Font.Ligatures Property (Word)

Returns or sets the ligatures setting for the specified  **Font** object. Read/write[WdLigatures](wdligatures-enumeration-word.md).


## Syntax

 _expression_ . **Ligatures**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Open Type fonts support the use of ligatures. The  **Ligatures** property specifies the Word ligatures setting to apply to an Open Type font.

The following table lists the four basic values for ligatures.



|**Value**|**Description**|
|:-----|:-----|
|Standard|Designed to enhance readability and attractiveness. Standard ligatures in Latin languages include "fi", "fl", and "ff".|
|Contextual|Designed to enhance readability and attractiveness by providing the best ligature choice given the surrounding text.|
|Historical|Older, ornamental ligatures that may look archaic to the modern reader. Not specifically designed for readability.|
|Discretional|Designed to be ornamental and not designed to be readable.|
 Combinations of these four basic values form the set of available values for the **Ligatures** property. This set of values is represented in the[WdLigatures](wdligatures-enumeration-word.md) enumeration.


## Example

The following code example applies Discretional ligatures to the font in the active document.


```vb
ActiveDocument.Range.Font.Ligatures = wdLigaturesDiscretional
```


## See also


#### Concepts


[Font Object](font-object-word.md)

