
# ShapeRange.GetTop Method (Publisher)

Returns the distance of the shape's or shape range's top edge from the top edge of the leftmost page in the current view as a  **Single** in the specified units.


## Syntax

 _expression_. **GetTop**( **_Unit_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Unit|Required| **PbUnitType**|The units in which to return the distance.|

### Return Value

Single


## Remarks

The Unit parameter can be one of the  **[PbUnitType](e14ef7b5-46c2-dec6-3af2-56da77ba5491.md)** constants declared in the Microsoft Publisher type library.

Use the  **[GetLeft](e8f28ab3-f9da-eae7-2a21-b8b2505e9b44.md)** method to return the distance of a shape's or shape range's left edge from the left edge of the leftmost page in the current view.


## Example

The following example displays the distances from the left and top edges of the leftmost page to the left and top edges of shape range consisting of all the shapes on the first page. The distances are expressed in inches (to the nearest hundredth).


```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Distance from left: " _ 
 &; Format(.GetLeft(Unit:=pbUnitInch), "0.00") _ 
 &; " in" &; vbCr _ 
 &; "Distance from top: " _ 
 &; Format(.GetTop(Unit:=pbUnitInch), "0.00") _ 
 &; " in" 
End With
```

