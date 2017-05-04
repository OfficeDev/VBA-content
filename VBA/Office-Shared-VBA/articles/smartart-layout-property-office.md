---
title: SmartArt.Layout Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArt.Layout
ms.assetid: 5aa76408-9c49-2430-eaea-8893a341b106
---


# SmartArt.Layout Property (Office)

Retrieves or sets the Smart Art layout associated with the Smart Art graphic. Read/write


## Syntax

 _expression_. **Layout**

 _expression_ An expression that returns a **SmartArt** object.


## Example

The following code sets the Smart Art layout.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also


#### Concepts


[SmartArt Object](smartart-object-office.md)
#### Other resources


[SmartArt Object Members](smartart-members-office.md)

