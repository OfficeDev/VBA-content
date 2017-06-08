---
title: Application.ActivePresentation Property (PowerPoint)
keywords: vbapp10.chm503001
f1_keywords:
- vbapp10.chm503001
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActivePresentation
ms.assetid: 55ff4906-09e5-2c5c-0ed7-5f7a767542f7
ms.date: 06/08/2017
---


# Application.ActivePresentation Property (PowerPoint)

Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the presentation open in the active window. Read-only.


## Syntax

 _expression_. **ActivePresentation**

 _expression_ A variable that represents an **Application** object.


### Return Value

Presentation


## Remarks

 If an embedded presentation is in-place active, the **ActivePresentation** property returns the embedded presentation.


## Example

This example saves the loaded presentation to the application folder in a file named "TestFile."


```
MyPath = Application.Path &; "\TestFile"

Application.ActivePresentation.SaveAs MyPath
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

