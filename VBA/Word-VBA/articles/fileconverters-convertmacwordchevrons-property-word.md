---
title: FileConverters.ConvertMacWordChevrons Property (Word)
keywords: vbawd10.chm161087490
f1_keywords:
- vbawd10.chm161087490
ms.prod: word
api_name:
- Word.FileConverters.ConvertMacWordChevrons
ms.assetid: c0a1f60c-f3aa-a091-2088-ff571f653ed1
ms.date: 06/08/2017
---


# FileConverters.ConvertMacWordChevrons Property (Word)

Controls whether text enclosed in chevron characters (« ») is converted to merge fields. Read/write  **Long** . .


## Syntax

 _expression_ . **ConvertMacWordChevrons**

 _expression_ A variable that represents a **[FileConverters](fileconverters-object-word.md)** collection.


## Remarks

The  **ConvertMacWordChevrons** property can be any **WdChevronConvertRule** constants.

Word for the Macintosh version 4.0 and 5.x documents use chevron characters to denote mail merge fields.


## Example

This example sets the  **ConvertMacWordChevrons** property to convert the text enclosed in chevrons to mail merge fields, and then it opens the document named "Mac Word Document."


```
FileConverters.ConvertMacWordChevrons = wdAlwaysConvert 
Documents.Open FileName:="C:\Documents\Mac Word Document"
```


## See also


#### Concepts


[FileConverters Collection Object](fileconverters-object-word.md)

