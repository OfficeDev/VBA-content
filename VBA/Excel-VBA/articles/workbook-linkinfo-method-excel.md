---
title: Workbook.LinkInfo Method (Excel)
keywords: vbaxl10.chm199108
f1_keywords:
- vbaxl10.chm199108
ms.prod: excel
api_name:
- Excel.Workbook.LinkInfo
ms.assetid: 016295a3-72c1-95b3-c259-8f286b12b73c
ms.date: 06/08/2017
---


# Workbook.LinkInfo Method (Excel)

Returns the link date and update status.


## Syntax

 _expression_ . **LinkInfo**( **_Name_** , **_LinkInfo_** , **_Type_** , **_EditionRef_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the link.|
| _LinkInfo_|Required| **[XlLinkInfo](xllinkinfo-enumeration-excel.md)**|The type of information to be returned.|
| _Type_|Optional| **Variant**|One of the constants of  **[XlLinkInfoType](xllinkinfotype-enumeration-excel.md)** specifying the type of link to return.|
| _EditionRef_|Optional| **Variant**|If the link is an edition, this argument specifies the edition reference as a string in R1C1 style. This argument is required if there's more than one publisher or subscriber with the same name in the workbook.|

### Return Value

Variant


## Example

This example displays a message box if the link is updated automatically.


```vb
If ActiveWorkbook.LinkInfo( _ 
 "Word.Document|Document1!'!DDE_LINK1", xlUpdateState, _ 
 xlOLELinks) = 1 Then 
 MsgBox "Link updates automatically" 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

