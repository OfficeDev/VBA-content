---
title: Item Method (VBA Add-In Object Model)
keywords: vbob6.chm104043
f1_keywords:
- vbob6.chm104043
ms.prod: office
ms.assetid: 46074a24-356c-f003-d8cd-67807bea1bcd
ms.date: 06/08/2017
---


# Item Method (VBA Add-In Object Model)



Returns the indexed member of a [collection](vbe-glossary.md).
 **Syntax**
 _object_**.Item(**_index_**)**
The  **Item** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _index_|Required. An expression that specifies the position of a member of the collection. If a [numeric expression](vbe-glossary.md),  _index_ must be a number from 1 to the value of the collection's **Count** property. If a[string expression](vbe-glossary.md),  _index_ must correspond to the _key_[argument](vbe-glossary.md) specified when the member was added to the collection.|
The following table lists the collections and their corresponding  _key_ arguments for use with the **Item** method. The string you pass to the **Item** method must match the collection's _key_ argument.


|**Collection**|**Key argument**|
|:-----|:-----|
|**Windows**|**Caption** property setting|
|**LinkedWindows**|**Caption** property setting|
|**CodePanes**|No unique string is associated with this collection.|
|**VBProjects**|**Name** property setting|
|**VBComponents**|**Name** property setting|
|**References**|**Name** property setting|
|**Properties**|**Name** property setting|
 **Remarks**
The  _index_ argument can be a numeric value or a string containing the title of the object.


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.



