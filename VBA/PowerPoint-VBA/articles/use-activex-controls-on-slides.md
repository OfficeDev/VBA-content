---
title: Use ActiveX Controls on Slides
keywords: vbapp10.chm5194029
f1_keywords:
- vbapp10.chm5194029
ms.prod: powerpoint
ms.assetid: c33a5c5a-b83d-3079-dc00-57b423697ea4
ms.date: 06/08/2017
---


# Use ActiveX Controls on Slides

You can add controls to your slides to provide a sophisticated way to exchange information with the user while a slide show is running. For example, you could use controls on slides to allow viewers of a show designed to be run in a kiosk a way to choose options and then run a custom show based on the viewer's choices.

For general information on adding and working with controls, see [How to: Use ActiveX Controls on Documents](use-activex-controls-on-documents.md) and[How to: Create Custom Dialog Boxes](create-custom-dialog-boxes.md).

Keep the following points in mind when you are working with controls on slides.


- A control on a slide is in design mode except when the slide show is running.
    
- If you want a control to appear on all slides in a presentation, add it to the slide master.
    
- The  **Me** keyword in an event procedure for a control on a slide refers to the slide, not the control.
    
Writing event code for controls on slides is very similar to writing event code for controls on forms. The following procedure sets the background for the slide the button named "cmdChangeColor" is on when the button is clicked.



```vb
Private Sub cmdChangeColor_Click()
    With Me
        .FollowMasterBackground = Not .FollowMasterBackground
        .Background.Fill.PresetGradient _
            msoGradientHorizontal, 1, msoGradientBrass
    End With
End Sub
```

You may want to use controls to provide your slide show with navigation tools that are more complex than those built into PowerPoint. For example, if you add two buttons named "cmdBack" and "cmdForward" to the slide master and write the following code behind them, all slides based on the master (and set to show master background graphics) will have these professional-looking navigation buttons that will be active during a slide show. 



```vb
Private Sub cmdBack_Click()
    Me.Parent.SlideShowWindow.View.Previous
End Sub

Private Sub cmdForward_Click()
    Me.Parent.SlideShowWindow.View.Next
End Sub
```

To work with all the ActiveX controls on a slide without affecting the other shapes on the slide, you can construct a  **ShapeRange** collection that contains only controls. You can then apply properties and methods to the entire collection or iterate through the collection to work with each control individually. The following example aligns all the controls on slide one in the active presentation and distributes them vertically.



```vb
With ActivePresentation.Slides(1).Shapes
    numShapes = .Count
    If numShapes > 1 Then
        numControls = 0
        ReDim ctrlArray(1 To numShapes)
        For i = 1 To numShapes
            If .Item(i).Type = msoOLEControlObject Then
                numControls = numControls + 1
                ctrlArray(numControls) = .Item(i).Name
            End If
        Next
        If numControls > 1 Then
            ReDim Preserve ctrlArray(1 To numControls)
            Set ctrlRange = .Range(ctrlArray)
            ctrlRange.Distribute msoDistributeVertically, True
            ctrlRange.Align msoAlignLefts, True
        End If
    End If
End With
```


