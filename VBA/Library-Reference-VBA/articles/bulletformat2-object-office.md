---
title: BulletFormat2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.BulletFormat2
ms.assetid: ad4c2a05-c34d-fbd4-6b12-3153b94d2c4e
---


# BulletFormat2 Object (Office)

Represents bullet formatting.


## Example

The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .TextFrame.TextRange.ParagraphFormat.BulletFormat2 
 .Visible = True 
 .RelativeSize = 1.25 
 .Character = 169 
 With .Font 
 .Color.RGB = RGB(255, 255, 0) 
 .Name = "Symbol" 
 End With 
 End With 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

