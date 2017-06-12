---
title: Section.Footers Property (Word)
keywords: vbawd10.chm156827770
f1_keywords:
- vbawd10.chm156827770
ms.prod: word
api_name:
- Word.Section.Footers
ms.assetid: 2aa522ae-fc34-eb75-790f-85a8182f76c2
ms.date: 06/08/2017
---


# Section.Footers Property (Word)

Returns a  **[HeadersFooters](headersfooters-object-word.md)** collection that represents the footers in the specified section. Read-only.


## Syntax

 _expression_ . **Footers**

 _expression_ A variable that represents a **[Section](section-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx). To return a  **HeadersFooters** collection that represents the headers for the specified section, use the **[Headers](section-headers-property-word.md)** property.


## Example

This example adds a right-aligned page number to the primary footer in the first section in the active document.


```vb
With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) 
 .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight 
End With
```


## See also


#### Concepts


[Section Object](section-object-word.md)

