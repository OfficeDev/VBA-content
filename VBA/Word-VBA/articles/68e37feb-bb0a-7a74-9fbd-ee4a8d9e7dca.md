
# Rows.DistanceRight Property (Word)

Returns or sets the distance (in points) between the document text and the right edge of the specified table. Read/write  **Single** .


## Syntax

 _expression_ . **DistanceRight**

 _expression_ A variable that represents a **[Rows](cd83d0ef-f743-1886-54de-497017c5f542.md)** collection.


## Remarks

This property doesn't have any effect if  **WrapAroundText** is **False** .


## Example

This example sets text to wrap around the first table in the active document and sets the distance for wrapped text to 20 points on all sides of the table.


```vb
With ActiveDocument.Tables(1).Rows 
 .WrapAroundText = True 
 .DistanceLeft = 20 
 .DistanceRight = 20 
 .DistanceTop = 20 
 .DistanceBottom = 20 
End With
```


## See also


#### Concepts


[Rows Collection Object](cd83d0ef-f743-1886-54de-497017c5f542.md)
#### Other resources


[Rows Object Members](161b0ab1-9763-3095-9152-07d6536c0fa4.md)
