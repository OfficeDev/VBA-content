---
title: Workbook.Signatures Property (Excel)
keywords: vbaxl10.chm199237
f1_keywords:
- vbaxl10.chm199237
ms.prod: excel
api_name:
- Excel.Workbook.Signatures
ms.assetid: b45f8036-c2d7-6113-e95c-ff78ee6a1f46
ms.date: 06/08/2017
---


# Workbook.Signatures Property (Excel)

Returns the digital signatures for a workbook. Read-only.


## Syntax

 _expression_ . **Signatures**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

To digitally sign Excel workbooks and verify other signatures in them, you will need the Microsoft CryptoAPI and a unique digital signature certificate. The CryptoAPI is installed with Microsoft Internet Explorer 4.01 or later. You can obtain a digital signature certificate from a certification authority.


## Example


```vb
Sub AddSignature() 
 ActiveWorkbook.Signatures.AddSignatureLine 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

