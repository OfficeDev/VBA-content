
# ShapeRange.IncrementTop Method (Excel)

 **Last modified:** July 28, 2015

Moves the specified shape vertically by the specified number of points.

## Syntax

 _expression_. **IncrementTop**( **_Increment_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Single**|Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## See also


#### Concepts


 [ShapeRange Object](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)
#### Other resources


 [ShapeRange Object Members](1d1950c5-32ac-dfc0-8c19-07159a29a2a0.md)
