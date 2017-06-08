---
title: Presentation.AddToFavorites Method (PowerPoint)
keywords: vbapp10.chm583031
f1_keywords:
- vbapp10.chm583031
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.AddToFavorites
ms.assetid: 5bdef3c1-fef2-a90b-d2be-f244e3ff1a64
ms.date: 06/08/2017
---


# Presentation.AddToFavorites Method (PowerPoint)

Adds a shortcut that represents the current selection in the specified presentation to the Windows Favorites folder.


## Syntax

 _expression_. **AddToFavorites**

 _expression_ A variable that represents a **Presentation** object.


## Remarks

The shortcut name is the display name of the document, if that's available; otherwise, the shortcut name is as calculated in HLINK.DLL.


## Example

This example adds a hyperlink to the active presentation to the Favorites folder in the Windows program folder.


```vb
Application.ActivePresentation.AddToFavorites
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

