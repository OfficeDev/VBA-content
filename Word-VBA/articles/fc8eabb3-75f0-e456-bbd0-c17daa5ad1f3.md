
# Application.PointsToPixels Method (Word)

 **Last modified:** July 28, 2015

Converts a measurement from points to pixels. Returns the converted measurement as a  **Single**.

## Syntax

 _expression_. **PointsToPixels**( **_Points_**,  **_fVertical_**)

 _expression_Required. A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Points|Required| **Single**|The point value to be converted to pixels.|
|fVertical|Optional| **Variant**| **True** to return the result as vertical pixels; **False** to return the result as horizontal pixels.|

### Return Value

Single


## Example

This example displays the height and width in pixels of an object measured in points.


```
MsgBox "180x120 points is equivalent to " _ 
 &amp; PointsToPixels(180, False) &amp; "x" _ 
 &amp; PointsToPixels(120, True) _ 
 &amp; " pixels on this display."
```


## See also


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
