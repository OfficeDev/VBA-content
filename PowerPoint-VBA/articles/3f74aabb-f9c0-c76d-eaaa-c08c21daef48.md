
# Series.Paste Method (PowerPoint)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a **[Series](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)** object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](e985978e-f0cf-b809-ebe1-f5504e9e8df6.md)** property to **xlMarkerStylePicture**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example pastes a picture from the Clipboard into series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Paste

    End If

End With


```


## See also


#### Concepts


[Series Object](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)
#### Other resources


[Series Object Members](f7e7168d-3c6f-20db-1e75-56a101c69a70.md)
