---
title: TextFrame.PathFormat Property (Word)
keywords: vbawd10.chm162665365
f1_keywords:
- vbawd10.chm162665365
ms.prod: word
api_name:
- Word.TextFrame.PathFormat
ms.assetid: 16d389c8-eda3-dec6-a40c-056e70f51dec
ms.date: 06/08/2017
---


# TextFrame.PathFormat Property (Word)

Returns or sets the path type for the specified text frame. Read/write  **MsoPathType** .


## Syntax

 _expression_ . **PathFormat**

 _expression_ A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Remarks

The value of the  **PathFormat** property can be one of the following **MsoPathType** constants:


- msoPathType1
    
- msoPathType2
    
- msoPathType3
    
- msoPathType4
    
- msoPathTypeMixed
    
- msoPathTypeNone
    


The value  **msoPathTypeMixed** cannot be set. Setting the value **msoPathTypeNone** removes any existing path.


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

