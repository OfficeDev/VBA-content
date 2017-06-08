---
title: Workbooks Object (Excel)
keywords: vbaxl10.chm202072
f1_keywords:
- vbaxl10.chm202072
ms.prod: excel
api_name:
- Excel.Workbooks
ms.assetid: f768da57-013a-e652-0f5d-60b03aa4240a
ms.date: 06/08/2017
---


# Workbooks Object (Excel)

A collection of all the  **[Workbook](workbook-object-excel.md)** objects that are currently open in the Microsoft Excel application.


## Remarks

For more information about using a single  **Workbook** object, see the[Workbook](workbook-object-excel.md) object.


## Example

Use the  **[Workbooks](application-workbooks-property-excel.md)** property to return the **Workbooks** collection. The following example closes all open workbooks.


```
Workbooks.Close
```

Use the  **[Add](workbooks-add-method-excel.md)** method to create a new, empty workbook and add it to the collection. The following example adds a new, empty workbook to Microsoft Excel.




```
Workbooks.Add
```

Use the  **[Open](workbooks-open-method-excel.md)** method to open a file. This creates a new workbook for the opened file. The following example opens the file Array.xls as a read-only workbook.




```
Workbooks.Open FileName:="Array.xls", ReadOnly:=True
```


## Methods



|**Name**|
|:-----|
|[Add](workbooks-add-method-excel.md)|
|[CanCheckOut](workbooks-cancheckout-method-excel.md)|
|[CheckOut](workbooks-checkout-method-excel.md)|
|[Close](workbooks-close-method-excel.md)|
|[Open](workbooks-open-method-excel.md)|
|[OpenDatabase](workbooks-opendatabase-method-excel.md)|
|[OpenText](workbooks-opentext-method-excel.md)|
|[OpenXML](workbooks-openxml-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](workbooks-application-property-excel.md)|
|[Count](workbooks-count-property-excel.md)|
|[Creator](workbooks-creator-property-excel.md)|
|[Item](workbooks-item-property-excel.md)|
|[Parent](workbooks-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
