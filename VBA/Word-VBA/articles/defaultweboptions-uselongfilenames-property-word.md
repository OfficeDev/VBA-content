---
title: DefaultWebOptions.UseLongFileNames Property (Word)
keywords: vbawd10.chm165871622
f1_keywords:
- vbawd10.chm165871622
ms.prod: word
api_name:
- Word.DefaultWebOptions.UseLongFileNames
ms.assetid: 7897cd7d-3815-8fc5-e752-0d93dd257915
ms.date: 06/08/2017
---


# DefaultWebOptions.UseLongFileNames Property (Word)

 **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **UseLongFileNames**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Remarks

If you don't use long file names and your document has supporting files, Microsoft Word automatically organizes those files in a separate folder. Otherwise, use the  **OrganizeInFolder** property to determine whether supporting files are organized in a separate folder.


## Example

This example disallows the use of long file names as the global default for the application.


```vb
Application.DefaultWebOptions.UseLongFileNames = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

