---
title: SpinButton.Orientation Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 923f3fe6-0740-9e42-ad63-ca29d04aad70
ms.date: 06/08/2017
---


# SpinButton.Orientation Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether the control is oriented vertically or horizontally. Read/write.


## Syntax

 _expression_. **Orientation**

 _expression_A variable that represents a  **SpinButton** object.


## Remarks

The settings for  **Orientation** are:



|**Value**|**Description**|
|:-----|:-----|
|-1|Automatically determines the orientation based upon the dimensions of the control (default).|
|0|Control is rendered vertically.|
|1|Control is rendered horizontally.|
If you specify automatic orientation, the height and width of the control determine whether it appears horizontally or vertically. For example, if the control is wider than it is tall, it appears horizontally; if it is taller than it is wide, the control appears vertically.


