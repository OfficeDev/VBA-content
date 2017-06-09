---
title: Application.SlideSelectionChanged Event (PowerPoint)
keywords: vbapp10.chm621016
f1_keywords:
- vbapp10.chm621016
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideSelectionChanged
ms.assetid: a7bbdc4c-31e3-2072-8590-bced8bff6517
ms.date: 06/08/2017
---


# Application.SlideSelectionChanged Event (PowerPoint)

Occurs at different times depending on the current view.


## Syntax

 _expression_. **SlideSelectionChanged**( **_SldRange_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SldRange_|Required|**SlideRange**|The selection of slides. In most cases this would be a single slide (for example, in Slide View you navigate to the next slide), but in some cases this could be multiple slides (for example, a marquee selection in Slide Sorter View).|

## Remarks

To access the  **Application** events, declare an **Application** variable in the General Declarations section of your code. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft PowerPoint **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).

The following table describes when the event occurs in each of the views. 



|**View**|**Description**|
|:-----|:-----|
|Normal, Master|Occurs when the slide in the slide pane changes.|
|Slide Sorter|Occurs when the selection changes.|
|Slide, Notes|Occurs when the slide changes.|
|Outline|Does not occur.|

## Example

This example displays a message every time a user selects a different slide. This example assumes that an  **Application** object called PPTApp has been declared by using the **WithEvents** keyword.


```vb
Private Sub PPTApp_SlideSelectionChanged(ByVal SldRange As SlideRange)

    MsgBox "Slide selection changed."

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

