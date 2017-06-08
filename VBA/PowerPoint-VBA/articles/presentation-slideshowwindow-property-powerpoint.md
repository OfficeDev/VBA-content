---
title: Presentation.SlideShowWindow Property (PowerPoint)
keywords: vbapp10.chm583047
f1_keywords:
- vbapp10.chm583047
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SlideShowWindow
ms.assetid: 9cef9c42-7a65-bd2e-3277-0145cd2cd3b9
ms.date: 06/08/2017
---


# Presentation.SlideShowWindow Property (PowerPoint)

Returns a  **[SlideShowWindow](slideshowwindow-object-powerpoint.md)** object that represents the slide show window in which the specified presentation is running. Read-only.


## Syntax

 _expression_. **SlideShowWindow**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

SlideShowWindow


## Remarks

You can use this property in conjunction with the  **Me** keyword and the **Parent** property to return the slide show window in which an ActiveX control event was fired, as shown in the example.


## Example

The following example shows the Click event procedures for buttons named "cmdBack" and "cmdForward". If you add these buttons to the slide master and add these event procedures to them, all slides based on the master (and set to show master background graphics) will have these navigation buttons that will be active during a slide show. The  **Me** keyword returns the **Master** object that represents the slide master that contains the control. If the control were on an individual slide, the **Me** keyword in an event procedure for that control would return a **Slide** object.


```vb
Private Sub cmdBack_Click()

    Me.Parent.SlideShowWindow.View.Previous

End Sub



Private Sub cmdForward_Click()

    Me.Parent.SlideShowWindow.View.Next

End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

