---
title: DoCmd.Hourglass Method (Access)
keywords: vbaac10.chm4155
f1_keywords:
- vbaac10.chm4155
ms.prod: access
api_name:
- Access.DoCmd.Hourglass
ms.assetid: e032e879-6ce4-982d-08cb-f9622c000b11
ms.date: 06/08/2017
---


# DoCmd.Hourglass Method (Access)

The  **Hourglass** method carries out the Hourglass action in Visual Basic.


## Syntax

 _expression_. **Hourglass**( ** _HourglassOn_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HourglassOn_|Required|**Variant**|Use  **True** (?1) to display the hourglass icon (or another icon you've chosen). Use **False** (0) to display the normal mouse pointer.|

## Remarks

You can use the  **Hourglass** method to change the mouse pointer to an image of an hourglass (or another icon you've chosen) while a procedure is running. This method can provide a visual indication that the procedure is running. This is especially useful when a procedure takes a long time to run.

You often use this method if you've turned echo off by using the  **Echo** method. When echo is off, Microsoft Access suspends screen updates until the macro is finished.

Microsoft Access automatically resets the  _HourglassOn_ argument to **False** when the procedure finishes running.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

