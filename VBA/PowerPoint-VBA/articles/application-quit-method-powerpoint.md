---
title: Application.Quit Method (PowerPoint)
keywords: vbapp10.chm502022
f1_keywords:
- vbapp10.chm502022
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Quit
ms.assetid: d7040179-ca03-563f-5bd9-80a5fd5e5d4b
ms.date: 06/08/2017
---


# Application.Quit Method (PowerPoint)

Quits Microsoft PowerPoint. This is equivalent to clicking the  **Office** button and then clicking **Exit PowerPoint**.


## Syntax

 _expression_. **Quit**

 _expression_ A variable that represents an **Application** object.


## Remarks

To avoid being prompted to save changes, use either the  **Save** or **SaveAs** method to save all open presentations before calling the **Quit** method.


## Example

This example saves all open presentations and then quits PowerPoint.


```vb
With Application

    For Each w In .Presentations

        w.Save

    Next w

    .Quit

End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

