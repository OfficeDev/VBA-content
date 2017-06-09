---
title: WebOptions.RelyOnCSS Property (Excel)
keywords: vbaxl10.chm662073
f1_keywords:
- vbaxl10.chm662073
ms.prod: excel
api_name:
- Excel.WebOptions.RelyOnCSS
ms.assetid: 7921e4d8-f27f-4e87-e64e-ae0f8c5869c3
ms.date: 06/08/2017
---


# WebOptions.RelyOnCSS Property (Excel)

 **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the **[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)** property. **False** if HTML <FONT> tags and cascading style sheets are used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnCSS**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

You should set this property to  **True** if your Web browser supports cascading style sheets, as this will give you more precise layout and formatting control on your Web page and make it look more like your document (as it appears in Microsoft Excel).


## Example

This example enables the use of cascading style sheets as the global default for the application.


```vb
ThisWorkbook.WebOptions.RelyOnCSS = True
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

