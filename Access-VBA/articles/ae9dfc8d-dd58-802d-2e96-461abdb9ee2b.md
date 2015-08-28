
# ListObject.ShowAutoFilter Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 Returns **Boolean** to indicate whether the AutoFilter will be displayed. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ShowAutoFilter**

 _expression_A variable that represents a  **ListObject** object.


## Remarks
<a name="sectionSection1"> </a>

 **ShowAutoFilter** property defaults to **True** for a new **ListObject** object.


## Example
<a name="sectionSection2"> </a>

The following example displays the setting of the  **ShowAutoFilter** property the default list in Sheet 1 of the active workbook.


```
 
 Dim wrksht As Worksheet 
 Dim oListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListCol = wrksht.ListObjects(1) 
 
 Debug.Print oListCol.ShowAutoFilter
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ListObject Object](46de6c4f-8ce0-0c7d-da59-6e52f5eab612.md)
#### Other resources


 [ListObject Object Members](d34f895c-cf60-f644-866b-7b757716e7a6.md)
