---
title: WebOptions.UseDefaultFolderSuffix Method (Word)
keywords: vbawd10.chm165937253
f1_keywords:
- vbawd10.chm165937253
ms.prod: word
api_name:
- Word.WebOptions.UseDefaultFolderSuffix
ms.assetid: f31703d4-0020-ec34-bc70-a737e978c666
ms.date: 06/08/2017
---


# WebOptions.UseDefaultFolderSuffix Method (Word)

Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.


## Syntax

 _expression_ . **UseDefaultFolderSuffix**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Remarks

Microsoft Word uses the folder suffix when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **UseLongFileNames** and **OrganizeInFolder** properties are set to **True** ).

The suffix appears in the folder name after the document name. For example, if the document is called "Doc1" and the language is English, the folder name is Doc1_files. The available folder suffixes are listed in the  **FolderSuffix** property topic.


## Example

This example sets the folder suffix for the active document to the default suffix.


```vb
ActiveDocument.WebOptions.UseDefaultFolderSuffix
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

