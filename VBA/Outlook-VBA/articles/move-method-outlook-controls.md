---
title: Move Method (Outlook Controls)
keywords: olfm10.chm2000320
f1_keywords:
- olfm10.chm2000320
ms.prod: outlook
ms.assetid: 9974e4bb-4b66-24f5-bf17-3e835863847f
ms.date: 06/08/2017
---


# Move Method (Outlook Controls)

Moves a control to the specified location.


## Syntax

 _expression_. **Move**( **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**,  **_Layout_**)

 _expression_A variable that represents an Outlook control object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Optional| **Variant**|Single-precision value, in points, indicating the horizontal coordinate for the left edge of the object.|
|Top|Optional| **Variant**|Single-precision value, in points, that specifies the vertical coordinate for the top edge of the object.|
|Width|Optional| **Variant**|Single-precision value, in points, indicating the width of the object.|
|Height|Optional| **Variant**|Single-precision value, in points, indicating the height of the object.|
|Layout|Optional| **Variant**|A Boolean value indicating whether the  **Layout** event is initiated for the control's parent following this move. **False** is the default value.|

## Remarks

The maximum and minimum values for the  _Left_,  _Top_,  _Width_, and  _Height_ arguments vary from one application to another.

You can move a control to a specific location relative to the edges of the form that contains the control.

You can use named arguments, or you can enter the arguments by position. If you use named arguments, you can list the arguments in any order. If not, you must enter the arguments in the order shown, using commas to indicate the relative position of arguments you do not specify. Any unspecified arguments remain unchanged.


