---
title: Workbook.HasVBProject Property (Excel)
keywords: vbaxl10.chm199250
f1_keywords:
- vbaxl10.chm199250
ms.prod: excel
api_name:
- Excel.Workbook.HasVBProject
ms.assetid: b4293266-40d9-a8a4-80ff-8b19ec7ed823
ms.date: 06/08/2017
---


# Workbook.HasVBProject Property (Excel)

Returns a  **Boolean** that represents whether a workbook has an attached Microsoft Visual Basic for Applications project. Read-only **Boolean** .


## Syntax

 _expression_ . **HasVBProject**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

This property is most useful in programmatically determining whether a workbook needs to be saved into a macro-enabled file format. If saved in another format, macros and code projects contained within the document may be lost.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

