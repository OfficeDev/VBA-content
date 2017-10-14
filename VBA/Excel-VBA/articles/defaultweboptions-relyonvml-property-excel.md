---
title: DefaultWebOptions.RelyOnVML Property (Excel)
keywords: vbaxl10.chm660081
f1_keywords:
- vbaxl10.chm660081
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.RelyOnVML
ms.assetid: 414b81f1-2549-f9b3-a5a0-b36dcbf8a6e4
ms.date: 06/08/2017
---


# DefaultWebOptions.RelyOnVML Property (Excel)

 **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnVML**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

You can reduce file sizes by not generating images for drawing objects, if your Web browser supports Vector Markup Language (VML). For example, Microsoft Internet Explorer 5 supports this feature, and you should set the  **RelyOnVML** property to **True** if you are targeting this browser. For browsers that do not support VML, the image will not appear when you view a Web page saved with this property enabled.

For example, you should not generate images if your Web page uses image files that you have generated earlier, and if the location where you save the document is different from the final location of the page on the Web server.


## Example

This example specifies that images are generated when saving the worksheet to a Web page.


```vb
Workbooks(1).DefaultWebOptions.RelyOnVML = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

