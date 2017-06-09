---
title: ZOrder Method (Outlook Controls)
keywords: olfm10.chm2000460
f1_keywords:
- olfm10.chm2000460
ms.prod: outlook
ms.assetid: 62bf7af1-8935-fd5e-da70-1b93408e015e
ms.date: 06/08/2017
---


# ZOrder Method (Outlook Controls)

Places the object at the front or back of the z-order.


## Syntax

 _expression_. **ZOrder**( **_zPosition_**)

 _expression_A variable that represents an Outlook control object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|zPosition|Optional| **Variant**|A control's position, front or back, in the container's z-order.|

## Remarks

The settings for  _zPosition_ are:



|**Value**|**Description**|
|:-----|:-----|
|0|Places the control at the front of the z-order. The control appears on top of other controls (default).|
|1|Places the control at the back of the z-order. The control appears underneath other controls.|
The z-order determines how windows and controls are stacked when they are presented to the user. Items at the back of the z-order are overlaid by closer items; items at the front of the z-order appear to be on top of items at the back. When the  _zPosition_ argument is omitted, the object is brought to the front.

In design mode, the  **Bring to Front** or **Send To Back** commands set the z-order. **Bring to Front** is equivalent to using the **ZOrder** method and putting the object at the front of the z-order. **Send To Back** is equivalent to using **ZOrder** and putting the object at the back of the z-order.

You can't Undo or Redo layering commands, such as  **Send To Back** or **Bring to Front**. For example, if you select an object and click  **Move Backward** on the shortcut menu, you won't be able to Undo or Redo that action.


