
# PageSetup.FooterMargin Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets the distance from the bottom of the page to the footer, in points. Read/write  **Double**.

## Syntax

 _expression_. **FooterMargin**

 _expression_A variable that represents a  **PageSetup** object.


## Example

This example sets the footer margin of Sheet1 to 0.5 inch.


```
Worksheets("Sheet1").PageSetup.FooterMargin = _ 
 Application.InchesToPoints(0.5)
```


## See also


#### Concepts


 [PageSetup Object](2fd22df9-5987-f723-04a9-9a3f2e84ac81.md)
#### Other resources


 [PageSetup Object Members](feabe079-cb03-f560-6032-88f5585ec8a8.md)
