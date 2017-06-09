---
title: Presentation.HandoutMaster Property (PowerPoint)
keywords: vbapp10.chm583010
f1_keywords:
- vbapp10.chm583010
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.HandoutMaster
ms.assetid: d80a8e51-61db-8da0-1fda-20a043e62569
ms.date: 06/08/2017
---


# Presentation.HandoutMaster Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the handout master. Read-only.


## Syntax

 _expression_. **HandoutMaster**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Master


## Example

This example sets the background pattern on the handout master in the active presentation.


```vb
Application.ActivePresentation.HandoutMaster.Background.Fill _
    .Patterned msoPatternDarkHorizontal
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

