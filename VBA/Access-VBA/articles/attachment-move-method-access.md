---
title: Attachment.Move Method (Access)
keywords: vbaac10.chm13976
f1_keywords:
- vbaac10.chm13976
ms.prod: access
api_name:
- Access.Attachment.Move
ms.assetid: cd807ce2-79b8-0873-c035-7927bc91967d
ms.date: 06/08/2017
---


# Attachment.Move Method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

 _expression_. **Move**( ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents an **Attachment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Variant**|The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.|
| _Top_|Optional|**Variant**|The screen position in twips for the top edge of the object relative to the top edge of the Microsoft Access window.|
| _Width_|Optional|**Variant**|The desired width in twips of the object.|
| _Height_|Optional|**Variant**|The desired height in twips of the object.|

## Remarks

Only the  _Left_ argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specify _Width_ without specifying _Left_ and _Top_. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

In Datasheet view or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

