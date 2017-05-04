---
title: SmartArtLayouts Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtLayouts
ms.assetid: 25e33439-fb5e-01d7-1b85-01884a42ba68
---


# SmartArtLayouts Object (Office)

Represents a collection of Smart Art layout diagrams.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[SmartArtLayouts Object Members](smartartlayouts-members-office.md)

