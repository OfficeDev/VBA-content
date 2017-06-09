---
title: DefaultWebOptions.AlwaysSaveInDefaultEncoding Property (Excel)
keywords: vbaxl10.chm660087
f1_keywords:
- vbaxl10.chm660087
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.AlwaysSaveInDefaultEncoding
ms.assetid: 7da4191e-4502-0005-1577-ac9a872f9cfa
ms.date: 06/08/2017
---


# DefaultWebOptions.AlwaysSaveInDefaultEncoding Property (Excel)

 **True** if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened. **False** if the original encoding of the file is used. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **AlwaysSaveInDefaultEncoding**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

The  **[Encoding](defaultweboptions-encoding-property-excel.md)** property can be used to set the default encoding.


## Example

This example sets the encoding to the default encoding. The encoding is used when you save the document as a Web page.


```vb
Application.DefaultWebOptions.AlwaysSaveInDefaultEncoding = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

