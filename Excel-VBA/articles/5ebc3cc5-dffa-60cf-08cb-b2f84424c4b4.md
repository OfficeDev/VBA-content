
# Workbook.ChangeHistoryDuration Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the number of days shown in the shared workbook's change history. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ChangeHistoryDuration**

 _expression_A variable that represents a  **Workbook** object.


## Remarks
<a name="sectionSection1"> </a>

Any changes in the change history older than the setting for this property are removed when the workbook is closed.


## Example
<a name="sectionSection2"> </a>

This example sets the number of days shown in the change history for the active workbook if change tracking is enabled. Any changes in the change history older than the setting for this property are removed when the workbook is closed.


```
With ActiveWorkbook 
 If .KeepChangeHistory Then 
 .ChangeHistoryDuration = 7 
 End If 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
