---
title: Application.Options Property (PowerPoint)
keywords: vbapp10.chm502054
f1_keywords:
- vbapp10.chm502054
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Options
ms.assetid: 4f890917-68bc-bb02-914d-52ea8a82bbcb
ms.date: 06/08/2017
---


# Application.Options Property (PowerPoint)

Returns an  **[Options](options-object-powerpoint.md)** object that represents application options in Microsoft PowerPoint.


## Syntax

 _expression_. **Options**

 _expression_ A variable that represents an **Application** object.


### Return Value

Options


## Example

Use the  **Options** property to return the **Options** object. The following example sets three application options for PowerPoint.


```vb
Sub TogglePasteOptionsButton()

    With Application.Options

        If .DisplayPasteOptions = False Then

            .DisplayPasteOptions = True

        End If

    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

