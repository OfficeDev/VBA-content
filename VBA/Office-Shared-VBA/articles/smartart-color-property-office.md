---
title: SmartArt.Color Property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Color
ms.assetid: 65105010-9780-1b99-ef23-b924300bfccb
ms.date: 06/08/2017
---


# SmartArt.Color Property (Office)

Retrieves or sets the Smart Art color style applied to the Smart Art graphic. Read/write


## Syntax

 _expression_. **Color**

 _expression_ An expression that returns a **SmartArt** object.


## Example

The following code sets the color scheme of the Smart Art diagram.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also


#### Concepts


[SmartArt Object](smartart-object-office.md)
#### Other resources


[SmartArt Object Members](smartart-members-office.md)

