---
title: WebOptions.UseLongFileNames Property (Excel)
keywords: vbaxl10.chm662075
f1_keywords:
- vbaxl10.chm662075
ms.prod: excel
api_name:
- Excel.WebOptions.UseLongFileNames
ms.assetid: f30c4954-d691-3a36-1540-f280eea370d8
ms.date: 06/08/2017
---


# WebOptions.UseLongFileNames Property (Excel)

 **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **UseLongFileNames**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

If you don't use long file names and your document has supporting files, Microsoft Excel automatically organizes those files in a separate folder. Otherwise, use the  **[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)** property to determine whether supporting files are organized in a separate folder.


## Example

This example disallows the use of long file names as the global default for the application.


```vb
Application.DefaultWebOptions.UseLongFileNames = False 

```


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

