
# Rows.DistributeHeight Method (Word)

 **Last modified:** July 28, 2015

Adjusts the height of the specified rows or cells so that they're equal.

## Syntax

 _expression_. **DistributeHeight**

 _expression_Required. A variable that represents a  ** [Rows](cd83d0ef-f743-1886-54de-497017c5f542.md)** collection.


## Example

This example adjusts the height of the rows in the first table in the active document so that they're equal.


```
ActiveDocument.Tables(1).Rows.DistributeHeight
```

This example adjusts the height of the first three rows in the first table so that they're equal.




```
Dim rngTemp As Range 
 
If ActiveDocument.Tables.Count >= 1 Then 
 Set rngTemp = ActiveDocument.Range(Start:=ActiveDocument _ 
 .Tables(1).Rows(1).Range.Start, _ 
 End:=ActiveDocument.Tables(1).Rows(3).Range.End) 
 rngTemp.Rows.DistributeHeight 
End If
```


## See also


#### Concepts


 [Rows Collection Object](cd83d0ef-f743-1886-54de-497017c5f542.md)
#### Other resources


 [Rows Object Members](161b0ab1-9763-3095-9152-07d6536c0fa4.md)
