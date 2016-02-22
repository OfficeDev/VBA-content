
# Series.Paste Method (Word)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **[Series](212c323f-8acb-2ba7-1359-ab0f43268e77.md)** object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](d9ba7847-2785-0f29-7e6e-d4bb2d62fc2f.md)** property to **xlMarkerStylePicture** .


## Example

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


[Series Object](212c323f-8acb-2ba7-1359-ab0f43268e77.md)
#### Other resources


[Series Object Members](0bc84851-3f0a-15e0-ae2b-c36215709220.md)
