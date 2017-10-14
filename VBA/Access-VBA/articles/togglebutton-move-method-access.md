---
title: ToggleButton.Move Method (Access)
keywords: vbaac10.chm11763
f1_keywords:
- vbaac10.chm11763
ms.prod: access
api_name:
- Access.ToggleButton.Move
ms.assetid: 98cdc2e0-63b7-ed59-0fca-3d4db5f1cf4b
ms.date: 06/08/2017
---


# ToggleButton.Move Method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

 _expression_. **Move**( ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents a **ToggleButton** object.


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

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

