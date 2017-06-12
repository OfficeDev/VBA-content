---
title: Application.SlideShowOnNext Event (PowerPoint)
keywords: vbapp10.chm621023
f1_keywords:
- vbapp10.chm621023
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowOnNext
ms.assetid: de72c6d6-0794-ad1d-5b25-478caaafd099
ms.date: 06/08/2017
---


# Application.SlideShowOnNext Event (PowerPoint)

Occurs when the user clicks  **Next** to move within the current slide.


## Syntax

 _expression_. **SlideShowOnNext**( **_Wn_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required|**SlideShowWindow**|The active slideshow window.|

## Remarks

To access  **Application** object events, declare a variable to represent the **Application** object in the **General Declarations** section of your code. Then set the variable equal to the **Application** object for which you want to access events. For more information about using events with the Microsoft PowerPoint **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).

The  **SlideShowOnNext** event does not fire when users click **Next** to move to the next slide, but rather only when they click **Next** to move within a given slide, for example to run the next animation on the slide.


## Example

This example displays a message every time a user clicks  **Next** to move with the current slide. The example assumes that you have already declared an **Application** object named _App_ in the **General Declarations** section of your code, using the **WithEvents** keyword.


```vb
Private Sub App_SlideShowOnNext(ByVal Wn As SlideShowWindow)



    Debug.Print "User clicked Next to move within the current slide."

        

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

