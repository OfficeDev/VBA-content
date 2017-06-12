---
title: TextFrame.AutoSize Property (Excel)
keywords: vbaxl10.chm644081
f1_keywords:
- vbaxl10.chm644081
ms.prod: excel
api_name:
- Excel.TextFrame.AutoSize
ms.assetid: bf434f76-5749-8163-f737-b3bd624092d5
ms.date: 06/08/2017
---


# TextFrame.AutoSize Property (Excel)

 **True** if the size of the specified object is changed automatically to fit text within its boundaries. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoSize**

 _expression_ A variable that represents a **TextFrame** object.


## Example

This example adjusts the size of the text frame on shape one to fit its text.


```vb
Worksheets(1).Shapes(1).TextFrame.AutoSize = True
```


## See also


#### Concepts


[TextFrame Object](textframe-object-excel.md)

