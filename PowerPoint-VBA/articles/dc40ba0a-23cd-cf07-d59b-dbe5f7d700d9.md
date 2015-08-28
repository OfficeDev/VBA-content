
# Pane.Application Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns an  ** [Application](978c2b99-4271-b953-4283-73b5f3d96f41.md)**object that represents the creator of the specified object.

## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **Pane** object.


### Return Value

Application


## Example

In this example, a  ** [Presentation](ec75cf52-69f8-d35b-0a26-4a8da8a9683f.md)**object is passed to the procedure. The procedure adds a slide to the presentation and then saves the presentation in the folder where Microsoft PowerPoint is running.


```
Sub AddAndSave(pptPres As Presentation)

    pptPres.Slides.Add 1, 1

    pptPres.SaveAs pptPres.Application.Path &amp; "\Added Slide"

End Sub
```

This example displays the name of the application that created each linked OLE object on slide one in the active presentation.




```
For Each shpOle In ActivePresentation.Slides(1).Shapes

    If shpOle.Type = msoLinkedOLEObject Then

        MsgBox shpOle.OLEFormat.Application.Name

    End If

Next
```


## See also


#### Concepts


 [Pane Object](27862fd6-897d-893d-d5a8-b1e40b1b9d48.md)
#### Other resources


 [Pane Object Members](d395cb24-e88f-5503-791b-83ecfaf10a7d.md)
