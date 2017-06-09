---
title: Hyperlink.EmailSubject Property (Excel)
keywords: vbaxl10.chm536083
f1_keywords:
- vbaxl10.chm536083
ms.prod: excel
api_name:
- Excel.Hyperlink.EmailSubject
ms.assetid: 3fe6d6a1-8184-8ef5-eb6e-b96ce9732dbd
ms.date: 06/08/2017
---


# Hyperlink.EmailSubject Property (Excel)

Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address. Read/write  **String** .


## Syntax

 _expression_ . **EmailSubject**

 _expression_ A variable that represents a **Hyperlink** object.


## Remarks

This property is usually used with e-mail hyperlinks.

The value of this property takes precedence over any e-mail subject line you have specified by using the  **[Address](hyperlink-address-property-excel.md)** property of the same **Hyperlink** object.


## Example

This example sets the e-mail subject line for the first hyperlink in the first worksheet.


```vb
Worksheets(1).Hyperlinks(1).EmailSubject = "Quote Request"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-excel.md)

