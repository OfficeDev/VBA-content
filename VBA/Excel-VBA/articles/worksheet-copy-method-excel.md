---
title: Worksheet.Copy Method (Excel)
keywords: vbaxl10.chm174074
f1_keywords:
- vbaxl10.chm174074
ms.prod: excel
api_name:
- Excel.Worksheet.Copy
ms.assetid: ace07575-34f4-a4ae-0922-a2671f2df1ba
ms.date: 06/08/2017
---


# Worksheet.Copy Method (Excel)

Copies the sheet to another location in the current workbook or a new workbook.


## Syntax

 _expression_ . **Copy**( **_Before_** , **_After_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the copied sheet will be placed. You cannot specify  _Before_ if you specify _After_.|
| _After_|Optional| **Variant**|The sheet after which the copied sheet will be placed. You cannot specify  _After_ if you specify _Before_.|

## Remarks

If you don't specify either  _Before_ or _After_, Microsoft Excel creates a new [workbook](workbook-object-excel.md) that contains the copied sheet object that contains the copied[Worksheet](worksheet-object-excel.md) object. The newly created workbook holds the[Application.ActiveWorkbook Property (Excel)](application-activeworkbook-property-excel.md) property and contains a single worksheet. The single worksheet retains the[Worksheet.Name Property (Excel)](worksheet-name-property-excel.md) and[Worksheet.CodeName Property (Excel)](worksheet-codename-property-excel.md) properties of the source worksheet. If the copied worksheet held a worksheet code sheet in a VBA project, that is also carried into the new workbook.

An array selection of multiple worksheets can be copied to a new blank [Workbook Object (Excel)](workbook-object-excel.md) object in a similar manner.


## Example

This example copies Sheet1, placing the copy after Sheet3.


```vb
Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")
```

This example first copies Sheet1 to a new blank workbook, then saves and closes the new workbook.




```vb
Worksheets("Sheet1").Copy
With ActiveWorkbook 
     .SaveAs Filename:=Environ("TEMP") &; "\New1.xlsx", FileFormat:=xlOpenXMLWorkbook
     .Close SaveChanges:=False
End With

```

This example copies worksheets Sheet1, Sheet2 and Sheet4 to a new blank workbook, then saves and closes the new workbook.




```vb
Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Copy
With ActiveWorkbook
     .SaveAs Filename:=Environ("TEMP") &; "\New3.xlsx", FileFormat:=xlOpenXMLWorkbook 
     .Close SaveChanges:=False 
End With 

```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

