
# PivotCache.RefreshOnFileOpen Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the PivotTable cache is automatically updated each time the workbook is opened. The default value is **False**. Read/write  **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RefreshOnFileOpen**

 _expression_A variable that represents a  **PivotCache** object.


## Remarks
<a name="sectionSection1"> </a>

Query tables and PivotTable reports are not automatically refreshed when you open the workbook by using the  ** [Open](1d1c3fca-ae1a-0a91-65a2-6f3f0fb308a0.md)** method in Visual Basic. Use the ** [Refresh](2833d199-342c-9e2e-d1f8-88c33a74bac6.md)**method to refresh the data after the workbook is open.


## Example
<a name="sectionSection2"> </a>

This example causes the PivotTable cache to automatically update each time the workbook is opened.


```
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PivotCache Object](c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439.md)
#### Other resources


 [PivotCache Object Members](113f1109-e1c9-2c6e-0581-9fba82f278dc.md)
