---
title: Ruler2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Ruler2
ms.assetid: a1632624-cdae-08db-4b5d-78311dbb224a
---


# Ruler2 Object (Office)

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.


## Remarks

Use the  **Ruler2** property of the **TextFrame2** object to return the **Ruler2** object that represents the ruler for the text in the specified shape. Use the **TabStops2** property to return the **TabStops2** object that contains the tab stops on the ruler. Use the **Levels** property to return the **RulerLevels2** object that contains the indentation settings for text outline levels.


## Example

The following example sets a left-aligned tab stop at 2 inches (144 Points) and sets a hanging indent for the text in object two on slide one in the active PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2).TextFrame2.Ruler2 
 .TabStops2.Add ppTabStopLeft, 144 
 .Levels(1).FirstMargin = 0 
 .Levels(1).LeftMargin = 36 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

