
# Axis.CategoryNames Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets all the category names as a text array for the specified axis. Read/write  **Variant**.

## Syntax

 _expression_. **CategoryNames**

 _expression_A variable that represents an  ** [Axis](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses an array to set individual category names for the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlCategory).CategoryNames = _

            Array ("1985", "1986", "1987", "1988", "1989")

    End If

End With
```


## See also


#### Concepts


 [Axis Object](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)
#### Other resources


 [Axis Object Members](6c4c7cca-d62e-a7c0-b724-30d1be8a44c9.md)
