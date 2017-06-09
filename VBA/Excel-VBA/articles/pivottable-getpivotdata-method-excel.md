---
title: PivotTable.GetPivotData Method (Excel)
keywords: vbaxl10.chm235139
f1_keywords:
- vbaxl10.chm235139
ms.prod: excel
api_name:
- Excel.PivotTable.GetPivotData
ms.assetid: 2d4600dd-6ca4-569a-6f93-79f6dbd43a09
ms.date: 06/08/2017
---


# PivotTable.GetPivotData Method (Excel)

Returns a  **[Range](range-object-excel.md)** object with information about a data item in a PivotTable report.


## Syntax

 _expression_ . **GetPivotData**( **_DataField_** , **_Field1_** , **_Item1_** , **_Field2_** , **_Item2_** , **_Field3_** , **_Item3_** , **_Field4_** , **_Item4_** , **_Field5_** , **_Item5_** , **_Field6_** , **_Item6_** , **_Field7_** , **_Item7_** , **_Field8_** , **_Item8_** , **_Field9_** , **_Item9_** , **_Field10_** , **_Item10_** , **_Field11_** , **_Item11_** , **_Field12_** , **_Item12_** , **_Field13_** , **_Item13_** , **_Field14_** , **_Item14_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataField_|Optional| **Variant**|The name of the field containing the data for the PivotTable.|
| _Field1_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item1_|Optional| **Variant**|The name of an item in  _Field1_.|
| _Field2_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item2_|Optional| **Variant**|The name of an item in  _Field2_.|
| _Field3_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item3_|Optional| **Variant**|The name of an item in  _Field3_.|
| _Field4_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item4_|Optional| **Variant**|The name of an item in  _Field4_.|
| _Field5_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item5_|Optional| **Variant**|The name of an item in  _Field5_.|
| _Field6_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item6_|Optional| **Variant**|The name of an item in  _Field6_.|
| _Field7_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item7_|Optional| **Variant**|The name of an item in  _Field7_.|
| _Field8_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item8_|Optional| **Variant**|The name of an item in  _Field8_.|
| _Field9_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item9_|Optional| **Variant**|The name of an item in  _Field9_.|
| _Field10_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item10_|Optional| **Variant**|The name of an item in  _Field10_.|
| _Field11_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item11_|Optional| **Variant**|The name of an item in  _Field11_.|
| _Field12_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item12_|Optional| **Variant**|The name of an item in  _Field12_.|
| _Field13_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item13_|Optional| **Variant**|The name of an item in  _Field13_.|
| _Field14_|Optional| **Variant**|The name of a column or row field in the PivotTable report.|
| _Item14_|Optional| **Variant**|The name of an item in  _Field14_.|

### Return Value

Range


## Example

In this example, Microsoft Excel returns the quantity of chairs in the warehouse to the user. This example assumes a PivotTable report exists on the active worksheet. Also, this example assumes that, in the report, the title of the data field is "Quantity", a field titled "Warehouse" exists, and a data item titled "Chairs" exists in the Warehouse field.


```vb
Sub UseGetPivotData() 
 
 Dim rngTableItem As Range 
 
 ' Get PivotData for the quantity of chairs in the warehouse. 
 Set rngTableItem = ActiveCell. _ 
 PivotTable.GetPivotData("Quantity", "Warehouse", "Chairs") 
 
 MsgBox "The quantity of chairs in the warehouse is: " &; rngTableItem.Value 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

