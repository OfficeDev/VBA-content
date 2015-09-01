
# Walls Object (Word)

 **Last modified:** July 28, 2015

Represents the walls of a 3-D chart. 

## Remarks

This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.


## Example

Use the  ** [Walls](f45ae75a-c96c-4441-af81-aedf23787194.md)** property to return the **Walls** object. The following example sets the pattern on the walls for the first chart in the active document. If the chart is not a 3-D chart, this example will fail.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Walls.Interior.Pattern = xlGray75 
 End If 
End With
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Walls Object Members](ff55b62c-e618-2e72-be85-fbe67cefc9ad.md)
