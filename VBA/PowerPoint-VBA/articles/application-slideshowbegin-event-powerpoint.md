---
title: Application.SlideShowBegin Event (PowerPoint)
keywords: vbapp10.chm621011
f1_keywords:
- vbapp10.chm621011
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowBegin
ms.assetid: f70ca9cb-11a7-2a81-19bb-36e0b0ca0b97
ms.date: 06/08/2017
---


# Application.SlideShowBegin Event (PowerPoint)

Occurs when you start a slide show.


## Syntax

 _expression_. **SlideShowBegin**( **_Wn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required|**SlideShowWindow**|The slide show window initialized prior to this event.|

## Remarks

Microsoft PowerPoint creates the slide show window and passes it to this event. If one slide show branches to another, the  **SlideShowBegin** event does not occur again when the second slide show begins.

For information about using events with the  **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).


## Example

This example adjusts the size and position of the slide show window and then reactivates it.


```vb
Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)

    With Wn

        .Height = 325

        .Width = 400

        .Left = 100

        .Activate

    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

