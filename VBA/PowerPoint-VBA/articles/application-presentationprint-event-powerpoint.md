---
title: Application.PresentationPrint Event (PowerPoint)
keywords: vbapp10.chm621015
f1_keywords:
- vbapp10.chm621015
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationPrint
ms.assetid: 41a420b7-c5db-7869-6763-da9cec710d83
ms.date: 06/08/2017
---


# Application.PresentationPrint Event (PowerPoint)

Occurs before a presentation is printed.


## Syntax

 _expression_. **PresentationPrint**( **_Pres_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation to be printed.|

## Remarks

For information about using events with the  **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).


## Example

This example sets the  **PrintHiddenSlides** property to **True** so that every time the active presentation is printed, the hidden slides are printed as well.


```vb
Private Sub App_PresentationPrint(ByVal Pres As Presentation)

    Pres.PrintOptions.PrintHiddenSlides = True

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

