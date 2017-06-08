---
title: DoCmd.GoToPage Method (Access)
keywords: vbaac10.chm4153
f1_keywords:
- vbaac10.chm4153
ms.prod: access
api_name:
- Access.DoCmd.GoToPage
ms.assetid: 37fe25b3-85b2-f681-acfd-96dab039e58f
ms.date: 06/08/2017
---


# DoCmd.GoToPage Method (Access)

Carries out the GoToPage action in Visual Basic. .


## Syntax

 _expression_. **GoToPage**( ** _PageNumber_**, ** _Right_**, ** _Down_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageNumber_|Required|**Variant**|A numeric expression that's a valid page number for the active form. If you leave this argument blank, the focus stays on the current page. You can use the  _right_ and _down_ arguments to display the part of the page you want to see.|
| _Right_|Optional|**Variant**|A numeric expression that's a valid horizontal offset for the page.|
| _Down_|Optional|**Variant**|A numeric expression that's a valid vertical offset for the page.|

### Return Value

Nothing


## Remarks

The units for the  _right_ and _down_ arguments are expressed in twips.

If you specify the  _right_ and _down_ arguments and leave the _pagenumber_ argument blank, you must include the _pagenumber_ argument's comma. If you don't specify the _right_ and _down_ arguments, don't use a comma following the _pagenumber_ argument.

The  **GoToPage** method of the **DoCmd** object was added to provide backwards compatibility for running the GoToPage action in Visual Basic code in Microsoft Access 95. It's recommended that you use the existing **GoToPage** method of the **Form** object instead.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

