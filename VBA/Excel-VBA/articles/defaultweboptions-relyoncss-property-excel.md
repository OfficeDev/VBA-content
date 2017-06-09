---
title: DefaultWebOptions.RelyOnCSS Property (Excel)
keywords: vbaxl10.chm660073
f1_keywords:
- vbaxl10.chm660073
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.RelyOnCSS
ms.assetid: 7700b648-9313-db23-bf27-5b73f21e5bce
ms.date: 06/08/2017
---


# DefaultWebOptions.RelyOnCSS Property (Excel)

 **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the **[OrganizeInFolder](defaultweboptions-organizeinfolder-property-excel.md)** property. **False** if HTML <FONT> tags and cascading style sheets are used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnCSS**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

You should set this property to  **True** if your Web browser supports cascading style sheets, as this will give you more precise layout and formatting control on your Web page and make it look more like your document (as it appears in Microsoft Excel).


## Example

This example enables the use of cascading style sheets as the global default for the application.


```vb
Application.DefaultWebOptions.RelyOnCSS = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

