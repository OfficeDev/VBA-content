
# HPageBreak.Location Property (Excel)

Returns or sets the cell (a **Range** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell. Read/write **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)**.


## Syntax

 _expression_ . **Location**

 _expression_ A variable that represents a **HPageBreak** object.


## Example

This example sets the horizontal page-break location. Note that you must be in Page Break Preview mode in order to set it.


```vb
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```
**Note:** The **Location** property can only be used to set the horizontal page-break location. In order to change the location of a **VPageBreak**, you must use [**VPageBreak.Dragoff**](93e169e8-e2d6-4cca-bd82-2d11fdc1ae4c.md).

## See also


#### Concepts


[HPageBreak Object](8fc96958-33ab-8251-f627-4769b5eab97f.md)
#### Other resources


[HPageBreak Object Members](32b561ff-a0cf-142b-0a46-c622a42b6125.md)
