---
title: Workbook.PersonalViewPrintSettings Property (Excel)
keywords: vbaxl10.chm199123
f1_keywords:
- vbaxl10.chm199123
ms.prod: excel
api_name:
- Excel.Workbook.PersonalViewPrintSettings
ms.assetid: 6e4a0a9c-4eb0-d8e1-e9ce-8e9e618996b4
ms.date: 06/08/2017
---


# Workbook.PersonalViewPrintSettings Property (Excel)

 **True** if print settings are included in the user's personal view of the shared workbook. Read-write **Boolean** .


## Syntax

 _expression_ . **PersonalViewPrintSettings**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example removes print settings and filter and sort settings from the user's personal view of workbook two.


```vb
With Workbooks(2) 
 .PersonalViewListSettings = False 
 .PersonalViewPrintSettings = False 
End With
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

