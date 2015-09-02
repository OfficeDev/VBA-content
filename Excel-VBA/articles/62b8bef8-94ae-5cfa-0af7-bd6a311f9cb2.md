
# FormatConditions.Item Method (Excel)

 **Last modified:** July 28, 2015

Returns a single object from a collection.

## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **FormatConditions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The name or index number for the object.|

### Return Value

An Object value that represents an object contained by the collection.


## Example

This example sets format properties for an existing conditional format for cells E1:E10.


```
With Worksheets(1).Range("e1:e10").FormatConditions.Item(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
End With
```


## See also


#### Concepts


 [FormatConditions Object](2486d4b4-605c-76d8-132a-694c0c600a81.md)
#### Other resources


 [FormatConditions Object Members](0e5a3774-fe65-597f-9b97-3bba637b55cc.md)
