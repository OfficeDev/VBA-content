---
title: Workbook.ReloadAs Method (Excel)
keywords: vbaxl10.chm199189
f1_keywords:
- vbaxl10.chm199189
ms.prod: excel
api_name:
- Excel.Workbook.ReloadAs
ms.assetid: ce6a9d1a-7945-3dca-ff2d-a42289c2ccf9
ms.date: 06/08/2017
---


# Workbook.ReloadAs Method (Excel)

Reloads a workbook based on an HTML document, using the specified document encoding.


## Syntax

 _expression_ . **ReloadAs**( **_Encoding_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Encoding_|Required| **[MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)**|The encoding that is to be applied to the workbook.|

## Remarks

Only  **msoEncoding** constants that are applicable to HTML work with the **ReloadAs** method.


## Example

This example reloads the first workbook, using Western document encoding.


```vb
Workbooks(1).ReloadAs Encoding:=msoEncodingWestern
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

