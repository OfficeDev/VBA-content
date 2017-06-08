---
title: SmartArtLayout Object (Office)
ms.prod: office
api_name:
- Office.SmartArtLayout
ms.assetid: f8d9db83-86f7-4830-096d-5d15368ab6b1
ms.date: 06/08/2017
---


# SmartArtLayout Object (Office)

Represents a Smart Art diagram.


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


[SmartArtLayout Object Members](smartartlayout-members-office.md)

