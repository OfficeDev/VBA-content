
# Axis.DisplayUnitLabel Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the  ** [DisplayUnitLabel](4dd4df7d-91c1-9136-2d5b-cdb0794a7716.md)**object for the specified axis. Returns  **null** if the ** [HasDisplayUnitLabel](adbbbb89-55af-12f5-ec67-1e88424f3d81.md)**property is set to  **False**. Read-only.

## Syntax

 _expression_. **DisplayUnitLabel**

 _expression_A variable that represents an  ** [Axis](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the label caption to "Millions" for the value axis of the first chart in the active document, and then it turns off automatic font scaling.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue).DisplayUnitLabel

            .Caption = "Millions"

            .AutoScaleFont = False

        End With

    End If

End With
```


## See also


#### Concepts


 [Axis Object](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)
#### Other resources


 [Axis Object Members](6c4c7cca-d62e-a7c0-b724-30d1be8a44c9.md)
