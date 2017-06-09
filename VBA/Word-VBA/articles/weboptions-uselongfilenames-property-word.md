---
title: WebOptions.UseLongFileNames Property (Word)
keywords: vbawd10.chm165937157
f1_keywords:
- vbawd10.chm165937157
ms.prod: word
api_name:
- Word.WebOptions.UseLongFileNames
ms.assetid: 25676029-e480-ac84-076a-95d3a41a800d
ms.date: 06/08/2017
---


# WebOptions.UseLongFileNames Property (Word)

 **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **UseLongFileNames**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Remarks

If you don't use long file names and your document has supporting files, Microsoft Word automatically organizes those files in a separate folder. Otherwise, use the  **[OrganizeInFolder](weboptions-organizeinfolder-property-word.md)** property to determine whether supporting files are organized in a separate folder.


## Example

This example disallows the use of long file names as the global default for the application.


```vb
Application.DefaultWebOptions.UseLongFileNames = False
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

