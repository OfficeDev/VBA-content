
# ListObject.ShowAutoFilter Property (Excel)

 Returns **Boolean** to indicate whether the AutoFilter will be displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowAutoFilter**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

 **ShowAutoFilter** property defaults to **True** for a new **ListObject** object.


## Example

The following example displays the setting of the  **ShowAutoFilter** property the default list in Sheet 1 of the active workbook.


```vb
 
 Dim wrksht As Worksheet 
 Dim oListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListCol = wrksht.ListObjects(1) 
 
 Debug.Print oListCol.ShowAutoFilter
```


## See also


#### Concepts


[ListObject Object](46de6c4f-8ce0-0c7d-da59-6e52f5eab612.md)
#### Other resources


[ListObject Object Members](d34f895c-cf60-f644-866b-7b757716e7a6.md)
