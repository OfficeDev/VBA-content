
# ShapeRange.VerticalFlip Property (Word)

 **Last modified:** July 28, 2015

 **True** if the specified shape is flipped around the vertical axis. Read-only **MsoTriState**.

## Syntax

 _expression_. **VerticalFlip**

 _expression_Required. A variable that represents a  ** [ShapeRange](7112acc0-e241-16ef-77bc-101b72d05af0.md)** object.


## Example

This example restores each shape on  _myDocument_ to its original state if it has been flipped horizontally or vertically.


```
For Each s In ActiveDocument.Range.ShapeRange 
 If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
 If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```


## See also


#### Concepts


 [ShapeRange Collection Object](7112acc0-e241-16ef-77bc-101b72d05af0.md)
#### Other resources


 [ShapeRange Object Members](eb882d13-d724-26e9-7e6d-2af55e42bba1.md)
