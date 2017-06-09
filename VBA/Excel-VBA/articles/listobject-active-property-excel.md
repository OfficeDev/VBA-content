---
title: ListObject.Active Property (Excel)
keywords: vbaxl10.chm734081
f1_keywords:
- vbaxl10.chm734081
ms.prod: excel
api_name:
- Excel.ListObject.Active
ms.assetid: abe995da-6471-e611-ee04-d24f8518327c
ms.date: 06/08/2017
---


# ListObject.Active Property (Excel)

 Returns a **Boolean** value indicating whether a **[ListObject](listobject-object-excel.md)** object in a worksheet is active—that is, whether the active cell is inside the range of the **ListObject** object. Read-only **Boolean** .


## Syntax

 _expression_ . **Active**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

Because there is no  **Activate** method for the **ListObject** object, you can activate a **ListObject** object only by activating a cell range within the list.


## Example

The following example activates the list in the default  **ListObject** object in the first worksheet of the active workbook. Invoking the **Activate** method of the **[Range](range-object-excel.md)** object without cell parameters activates the entire range for the list.


```vb
Function MakeListActive() As Boolean 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objList = wrksht.ListObjects(1) 
 objList.Range.Activate 
 
 MakeListActive = objList.Active 
End Function
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

