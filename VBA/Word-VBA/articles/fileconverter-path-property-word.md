---
title: FileConverter.Path Property (Word)
keywords: vbawd10.chm161021958
f1_keywords:
- vbawd10.chm161021958
ms.prod: word
api_name:
- Word.FileConverter.Path
ms.assetid: 85809cfe-7db5-cada-9b25-3d6276356ea9
ms.date: 06/08/2017
---


# FileConverter.Path Property (Word)

Returns the disk or Web path to the specified object. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[Name](fileconverter-name-property-word.md)** property to return the file name without the path. You can create the full name of a file converter by concatenating the **Path** , **PathSeparator** , and **Name** properties.


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

