---
title: Application.Path Property (PowerPoint)
keywords: vbapp10.chm502008
f1_keywords:
- vbapp10.chm502008
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Path
ms.assetid: aae10b96-e0e4-d055-f398-d26f4cab572d
ms.date: 06/08/2017
---


# Application.Path Property (PowerPoint)

Returns a  **String** that represents the path to the specified **[Application](application-object-powerpoint.md)** object. Read-only.


## Syntax

 _expression_. **Path**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

The path doesn't include the final backslash (\) or the name of the specified object. Use the  **Name** property of the **Presentation** object to return the file name without the path, and use the **FullName** property to return the file name and the path together.


## Example

This example saves the active presentation in the same folder as PowerPoint. 


```vb
With Application

    fName = .Path &; "\test presentation"

    ActivePresentation.SaveAs fName

End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

