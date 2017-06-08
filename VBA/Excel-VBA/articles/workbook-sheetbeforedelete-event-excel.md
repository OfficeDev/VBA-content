---
title: Workbook.SheetBeforeDelete Event (Excel)
keywords: vbaxl10.chm503112
f1_keywords:
- vbaxl10.chm503112
ms.assetid: 42406738-0fcd-4ef7-9bd6-abcc05f5e922
ms.date: 06/08/2017
ms.prod: excel
---


# Workbook.SheetBeforeDelete Event (Excel)

Occurs when any sheet is deleted.


## Syntax

 _expression_ . **SheetBeforeDelete**( **_Sh_** , )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The sheet. Can be a  **[Chart](chart-object-excel.md)** or **[Worksheet](worksheet-object-excel.md)** object.|

## Example

This example displays the name of each deactivated sheet.


```vb
Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

