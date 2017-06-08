---
title: Move Method
keywords: fm20.chm5224963
f1_keywords:
- fm20.chm5224963
ms.prod: office
api_name:
- Office.Move
ms.assetid: b4138364-0f1a-b774-a82b-3417cb9a6599
ms.date: 06/08/2017
---


# Move Method



Moves a form or control, or moves all the controls in the  **Controls** collection.
 **Syntax**
For a form or control _object_. **Move(** [ _Left_ [, _Top_ [, _Width_ [, _Height_ [, _Layout_ ]]]]] **)**
For the Controls collection _object_. **Move(** X, Y **)**
The  **Move** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _Left_|Optional. [Single-precision value](vbe-glossary.md), in points, indicating the horizontal coordinate for the left edge of the object.|
| _Top_|Optional. Single-precision value, in points, that specifies the vertical coordinate for the top edge of the object.|
| _Width_|Optional. Single-precision value, in points, indicating the width of the object.|
| _Height_|Optional. Single-precision value, in points, indicating the height of the object.|
| _Layout_|Optional. A Boolean value indicating whether the Layout event is initiated for the control's parent following this move.  **False** is the default value.|
| _X, Y_|Required. Single-precision value, in points, that specifies the change from the current horizontal and vertical position for each control in the  **Controls** collection.|
 **Settings**
The maximum and minimum values for the  _Left_, _Top_, _Width_, _Height_, _X_, and _Y_ arguments vary from one application to another.
 **Remarks**
For a form or control, you can move a selection to a specific location relative to the edges of the form that contains the selection.
You can use [named arguments](vbe-glossary.md), or you can enter the arguments by position. If you use named arguments, you can list the arguments in any order. If not, you must enter the arguments in the order shown, using commas to indicate the relative position of arguments you do not specify. Any unspecified arguments remain unchanged.
For the  **Controls** collection, you can move all the controls in this collection a specific distance from their current positions on a form, **Frame**, or **Page**.

