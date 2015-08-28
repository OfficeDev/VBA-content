
# HPageBreak.Location Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets the cell (a  **Range** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/write ** [Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)**.

## Syntax

 _expression_. **Location**

 _expression_A variable that represents a  **HPageBreak** object.


## Example

This example moves the horizontal page-break location.


```
Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```


## See also


#### Concepts


 [HPageBreak Object](8fc96958-33ab-8251-f627-4769b5eab97f.md)
#### Other resources


 [HPageBreak Object Members](32b561ff-a0cf-142b-0a46-c622a42b6125.md)
