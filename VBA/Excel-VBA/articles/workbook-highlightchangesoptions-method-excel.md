---
title: Workbook.HighlightChangesOptions Method (Excel)
keywords: vbaxl10.chm199172
f1_keywords:
- vbaxl10.chm199172
ms.prod: excel
api_name:
- Excel.Workbook.HighlightChangesOptions
ms.assetid: ac69ee3e-c5ea-5ac0-418a-0b94d56a8777
ms.date: 06/08/2017
---


# Workbook.HighlightChangesOptions Method (Excel)

Controls how changes are shown in a shared workbook.


## Syntax

 _expression_ . **HighlightChangesOptions**( **_When_** , **_Who_** , **_Where_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _When_|Optional| **Variant**|The changes that are shown. Can be one of the following  **[XlHighlightChangesTime](xlhighlightchangestime-enumeration-excel.md)** constants: **xlSinceMyLastSave** , **xlAllChanges** , or **xlNotYetReviewed** .|
| _Who_|Optional| **Variant**|The user or users whose changes are shown. Can be "Everyone," "Everyone but Me," or the name of one of the users of the shared workbook.|
| _Where_|Optional| **Variant**|An A1-style range reference that specifies the area to check for changes.|

## Example

This example shows changes to the shared workbook on a separate worksheet.


```vb
With ActiveWorkbook 
 .HighlightChangesOptions _ 
 When:=xlSinceMyLastSave, _ 
 Who:="Everyone" 
 .ListChangesOnNewSheet = True 
End With 

```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

