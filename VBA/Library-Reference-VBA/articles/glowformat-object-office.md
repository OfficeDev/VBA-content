---
title: GlowFormat Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.GlowFormat
ms.assetid: b89e2245-e3a4-4a8c-cd4f-86396ad71a5b
---


# GlowFormat Object (Office)

Represents a glow effect around an Office graphic.


## Example

This example applies glow to the text in the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(2).Shapes(2) 
 .Text.Font.Glowformat = msoGlowType2 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

