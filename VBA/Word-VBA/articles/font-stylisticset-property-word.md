---
title: Font.StylisticSet Property (Word)
keywords: vbawd10.chm156369074
f1_keywords:
- vbawd10.chm156369074
ms.prod: word
api_name:
- Word.Font.StylisticSet
ms.assetid: e82013b1-9f55-d17a-a510-6f77b627382b
ms.date: 06/08/2017
---


# Font.StylisticSet Property (Word)

Specifies the stylistic set for the specified font. Read/write [WdStylisticSet](wdstylisticset-enumeration-word.md).


## Syntax

 _expression_ . **StylisticSet**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Some OpenType fonts provide stylistic sets. A stylistic set defines a set of characters within the font that are intended to be used together, usually for the purpose of visual harmony, such as in headings.


## Example

The following code example sets the font for the active document to Gabriola and then applies the sixth stylistic set provided by the Gabriola font.


```vb
ActiveDocument.Range.Font.Name = "Gabriola" 
ActiveDocument.Range.Font.StylisticSet = wdStylisticSet06
```


## See also


#### Concepts


[Font Object](font-object-word.md)

