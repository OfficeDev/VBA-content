---
title: DefaultWebOptions.OrganizeInFolder Property (Word)
keywords: vbawd10.chm165871620
f1_keywords:
- vbawd10.chm165871620
ms.prod: word
api_name:
- Word.DefaultWebOptions.OrganizeInFolder
ms.assetid: 318d8f6d-79c6-9ea1-dc17-d605ce184d75
ms.date: 06/08/2017
---


# DefaultWebOptions.OrganizeInFolder Property (Word)

 **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **OrganizeInFolder**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Remarks

The new folder is created in the folder where you have saved the Web page and is named after the document. If long file names are used, a suffix is added to the folder name. The  **FolderSuffix** property returns wither the folder suffix for the language support you have selected or installed or the default folder suffix.

If you save a document that was previously saved with the  **OrganizeInFolder** property set to a different value, Microsoft Word automatically moves the supporting files into or out of the folder, as appropriate.

If you don't use long file names (that is, if the  **UseLongFileNames** property is set to **False** ), Microsoft Word automatically saves any supporting files in a separate folder. The files cannot be saved in the same folder as the Web page.


## Example

This example specifies that all supporting files are saved in the same folder when the document is saved as a Web page.


```vb
Application.DefaultWebOptions.OrganizeInFolder = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

