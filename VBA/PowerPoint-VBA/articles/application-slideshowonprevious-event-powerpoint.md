---
title: Application.SlideShowOnPrevious Event (PowerPoint)
keywords: vbapp10.chm621024
f1_keywords:
- vbapp10.chm621024
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowOnPrevious
ms.assetid: 466a5363-047b-f107-011b-6450db6a5f31
ms.date: 06/08/2017
---


# Application.SlideShowOnPrevious Event (PowerPoint)

Occurs when the user clicks  **Previous** to move within the current slide.


## Syntax

 _expression_. **SlideShowOnPrevious**( **_Wn_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required|**SlideShowWindow**|The active slideshow window.|

## Remarks

To access  **Application** object events, declare a variable to represent the **Application** object in the **General Declarations** section of your code. Then set the variable equal to the **Application** object for which you want to access events. For more information about using events with the Microsoft PowerPoint **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).

The  **SlideShowOnPrevious** event does not fire when users click **Previous** to move from one slide to the previous one, but rather only when they click **Previous** to move within a given slide, for example to rerun the previous animation on the slide.


## Example

This example displays a message every time a user clicks  **Previous** to move with the current slide. The example assumes that you have already declared an **Application** object named _App_ in the **General Declarations** section of your code, using the **WithEvents** keyword.


```vb
Private Sub App_SlideShowOnPrevious(ByVal Wn As SlideShowWindow)



    Debug.Print "User clicked Previous to move within the current slide."

        

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

