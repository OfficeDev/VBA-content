---
title: DefaultWebOptions.RelyOnVML Property (Word)
keywords: vbawd10.chm165871625
f1_keywords:
- vbawd10.chm165871625
ms.prod: word
api_name:
- Word.DefaultWebOptions.RelyOnVML
ms.assetid: b062a449-11f3-3467-994b-d854f85d064f
ms.date: 06/08/2017
---


# DefaultWebOptions.RelyOnVML Property (Word)

 **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnVML**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Remarks

You can reduce file sizes by not generating images for drawing objects, if your Web browser supports Vector Markup Language (VML). For example, Microsoft Internet Explorer 5 supports this feature, and you should set the  **RelyOnVML** property to **True** if you are targeting this browser. For browsers that do not support VML, the image will not appear when you view a Web page saved with this property enabled.

Don't generate images if your Web page uses image files that you have generated earlier and if the location where you save the document is different from the final location of the page on the Web server.


## Example

This example specifies that images are generated when saving the document as a Web page.


```vb
Application.DefaultWebOptions.RelyOnVML = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

