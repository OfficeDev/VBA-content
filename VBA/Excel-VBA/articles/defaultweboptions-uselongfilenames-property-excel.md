---
title: DefaultWebOptions.UseLongFileNames Property (Excel)
keywords: vbaxl10.chm660078
f1_keywords:
- vbaxl10.chm660078
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.UseLongFileNames
ms.assetid: b594ad04-866a-b811-338b-73d45352866b
ms.date: 06/08/2017
---


# DefaultWebOptions.UseLongFileNames Property (Excel)

 **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **UseLongFileNames**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

If you don't use long file names and your document has supporting files, Microsoft Excel automatically organizes those files in a separate folder. Otherwise, use the  **[OrganizeInFolder](defaultweboptions-organizeinfolder-property-excel.md)** property to determine whether supporting files are organized in a separate folder.


## Example

This example disallows the use of long file names as the global default for the application.


```vb
Application.DefaultWebOptions.UseLongFileNames = False 

```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

