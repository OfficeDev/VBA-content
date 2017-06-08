---
title: DoCmd.MoveSize Method (Access)
keywords: vbaac10.chm4158
f1_keywords:
- vbaac10.chm4158
ms.prod: access
api_name:
- Access.DoCmd.MoveSize
ms.assetid: 8fe8fc60-023e-26ce-c11a-2c29ffc21fbb
ms.date: 06/08/2017
---


# DoCmd.MoveSize Method (Access)

The  **MoveSize** method carries out the MoveSize action in Visual Basic.


## Syntax

 _expression_. **MoveSize**( ** _Right_**, ** _Down_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Right_|Optional|**Variant**|The new horizontal position of the window's upper-left corner, measured from the left edge of its containing window.|
| _Down_|Optional|**Variant**|The new vertical position of the window's upper-left corner, measured from the top edge of its containing window.|
| _Width_|Optional|**Variant**|The window's new width.|
| _Height_|Optional|**Variant**|The window's new height.|

## Remarks

You can use the  **MoveSize** method to move or resize the active window.

The units for the arguments are twips.

You must include at least one argument for the  **MoveSize** method. If you leave an argument blank, the current setting for the window is used.


## Example

The following example moves the active window and changes its height, but leaves its width unchanged:


```vb
DoCmd.MoveSize 1440, 2400, , 2000
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

