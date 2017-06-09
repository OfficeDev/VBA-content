---
title: Application.SlideShowNextSlide Event (PowerPoint)
keywords: vbapp10.chm621013
f1_keywords:
- vbapp10.chm621013
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowNextSlide
ms.assetid: a73d051e-9f53-43bd-1f41-b9111197e464
ms.date: 06/08/2017
---


# Application.SlideShowNextSlide Event (PowerPoint)

Occurs immediately before the transition to the next slide. For the first slide, occurs immediately after the  **[SlideShowBegin](application-slideshowbegin-event-powerpoint.md)** event.


## Syntax

 _expression_. **SlideShowNextSlide**( **_Wn_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required|**SlideShowWindow**|The active slide show window.|

## Remarks

For information about using events with the  **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).


## Example

This example determines the slide position for the slide following the  **SlideShowNextSlide** event. If the next slide is slide three, the example changes the type of pointer to a pen and the pen color to red.


```vb
Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)



   Dim Showpos As Integer



   Showpos = Wn.View.CurrentShowPosition + 1

If Showpos = 3 Then  

         With ActivePresentation.SlideShowSettings.Run.View

            .PointerColor.RGB = RGB(255, 0, 0)

            .PointerType = ppSlideShowPointerPen

         End With

      Else

         With ActivePresentation.SlideShowSettings.Run.View

            .PointerColor.RGB = RGB(0, 0, 0)

            .PointerType = ppSlideShowPointerArrow

         End With

      End If

End Sub
```

This example sets a global counter variable to zero. Then it calculates the number of shapes on the slide following this event, determines which shapes have animation, and fills a global array with the animation order and the number of each shape.


 **Note**  The array created in this example is also used in the  **SlideShowNextBuild** event example.




```vb
Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)



   Dim i as Integer, j as Integer, numShapes As Integer

   Dim objSld As Slide



   Set objSld = ActivePresentation.Slides _
        (ActivePresentation.SlideShowWindow.View _
        .CurrentShowPosition + 1)

      With objSld.Shapes

         numShapes = .Count

         If numShapes > 0 Then

            j = 1

            ReDim shpAnimArray(1 To 2, 1 To numShapes)

            For i = 1 To numShapes

               If .Item(i).AnimationSettings.Animate Then

                  shpAnimArray(1, j) = _
                     .Item(i).AnimationSettings.AnimationOrder

                     shpAnimArray(2, j) = i

                     j = j + 1

               End If

            Next

         End If

      End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

