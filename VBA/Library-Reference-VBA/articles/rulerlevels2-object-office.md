---
title: RulerLevels2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.RulerLevels2
ms.assetid: 01bd257c-1c26-a7cd-cf2a-8478c861b78a
---


# RulerLevels2 Object (Office)

A collection of all the  **RulerLevel2** objects on the specified ruler.


## Remarks

Each  **RulerLevel2** object represents the first-line and left indent for text at a particular outline level. This collection always contains five members — one for each of the available outline levels.


## Example

Use the  **Levels** property to return the RulerLevels2 collection. The following example sets the margins for the five outline levels in body text in the active presentation.


```vb
With ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Ruler2 
 .Levels(1).FirstMargin = 0 
 .Levels(1).LeftMargin = 40 
 .Levels(2).FirstMargin = 60 
 .Levels(2).LeftMargin = 100 
 .Levels(3).FirstMargin = 120 
 .Levels(3).LeftMargin = 160 
 .Levels(4).FirstMargin = 180 
 .Levels(4).LeftMargin = 220 
 .Levels(5).FirstMargin = 240 
 .Levels(5).LeftMargin = 280 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

