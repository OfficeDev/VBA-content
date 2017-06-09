---
title: PageSetup.SuppressEndnotes Property (Word)
keywords: vbawd10.chm158400629
f1_keywords:
- vbawd10.chm158400629
ms.prod: word
api_name:
- Word.PageSetup.SuppressEndnotes
ms.assetid: be1a8712-8763-646f-6126-30fa0056f159
ms.date: 06/08/2017
---


# PageSetup.SuppressEndnotes Property (Word)

 **True** if endnotes are printed at the end of the next section that doesn't suppress endnotes. Read/write **Long** .


## Syntax

 _expression_ . **SuppressEndnotes**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

Suppressed endnotes are printed before the endnotes in that section. This property takes effect only if the  **[Location](endnotes-location-property-word.md)** property is set to **wdEndOfSection** .


## Example

This example suppresses endnotes in the first section of the active document.


```vb
If ActiveDocument.Endnotes.Location = wdEndOfSection Then 
 ActiveDocument.Sections(1).PageSetup.SuppressEndnotes = True 
End If
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

