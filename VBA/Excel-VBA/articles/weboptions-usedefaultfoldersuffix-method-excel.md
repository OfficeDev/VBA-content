---
title: WebOptions.UseDefaultFolderSuffix Method (Excel)
keywords: vbaxl10.chm662084
f1_keywords:
- vbaxl10.chm662084
ms.prod: excel
api_name:
- Excel.WebOptions.UseDefaultFolderSuffix
ms.assetid: dbaf5fa4-449a-b549-d2a0-82f65497f6c0
ms.date: 06/08/2017
---


# WebOptions.UseDefaultFolderSuffix Method (Excel)

Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.


## Syntax

 _expression_ . **UseDefaultFolderSuffix**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

Microsoft Excel uses the folder suffix when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **[UseLongFileNames](weboptions-uselongfilenames-property-excel.md)** and **[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)** properties are set to **True** ).

The suffix appears in the folder name after the document name. For example, if the document is called "Book1" and the language is English, the folder name is Book1_files. The available folder suffixes are listed in the  **[FolderSuffix](weboptions-foldersuffix-property-excel.md)** property topic.


## Example

This example sets the folder suffix for the first workbook to the default suffix.


```vb
Workbooks(1).WebOptions.UseDefaultFolderSuffix
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

