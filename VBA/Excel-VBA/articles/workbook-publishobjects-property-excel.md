---
title: Workbook.PublishObjects Property (Excel)
keywords: vbaxl10.chm199187
f1_keywords:
- vbaxl10.chm199187
ms.prod: excel
api_name:
- Excel.Workbook.PublishObjects
ms.assetid: b6418f80-5154-6e3f-7313-222e6438c0e1
ms.date: 06/08/2017
---


# Workbook.PublishObjects Property (Excel)

Returns the  **[PublishObjects](publishobjects-object-excel.md)** collection. Read-only.


## Syntax

 _expression_ . **PublishObjects**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example publishes all static  **PublishObject** objects in the active workbook to the Web page.


```vb
Set objPObjs = ActiveWorkbook.PublishObjects 
For Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

