---
title: Template.Path Property (Word)
keywords: vbawd10.chm157941761
f1_keywords:
- vbawd10.chm157941761
ms.prod: word
api_name:
- Word.Template.Path
ms.assetid: 9b84e053-b806-d43d-2c3c-b8ce56cf7d15
ms.date: 06/08/2017
---


# Template.Path Property (Word)

Returns the path to the specified document template. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[Name](template-name-property-word.md)** property to return the file name without the path and use the **[FullName](template-fullname-property-word.md)** property to return the file name and the path together.


## See also


#### Concepts


[Template Object](template-object-word.md)

