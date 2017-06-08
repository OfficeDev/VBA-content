---
title: Form.InsideWidth Property (Access)
keywords: vbaac10.chm13476
f1_keywords:
- vbaac10.chm13476
ms.prod: access
api_name:
- Access.Form.InsideWidth
ms.assetid: c92954cd-0b8b-94d8-8826-684e886da0a2
ms.date: 06/08/2017
---


# Form.InsideWidth Property (Access)

You can use the  **InsideWidth** property (along with the **InsideHeight** property) to determine the height and width (in twips) of the window containing a form. Read/write **Long**.


## Syntax

 _expression_. **InsideWidth**

 _expression_ A variable that represents a **Form** object.


## Remarks

If you want to determine the interior dimensions of the form itself, you use the  **Width** property to determine the form width and the sum of the heights of the form's visible sections to determine its height (the **Height** property applies only to form sections, not to forms). The interior of a form is the region inside the form, excluding the scroll bars and the record selectors.

You can also use the  **WindowHeight** and **WindowWidth** properties to determine the height and width of the window containing a form.

If a window is maximized, setting these properties doesn't have any effect until the window is restored to its normal size.


## Example

The following example shows how to use the  **InsideHeight** and **InsideWidth** properties to compare the inside height and width of a form with the height and width of the form's window. If the window's dimensions don't equal the size of the form, then the window is resized to match the form's height and width.


```vb
Sub ResetWindowSize(frm As Form) 
 Dim intWindowHeight As Integer 
 Dim intWindowWidth As Integer 
 Dim intTotalFormHeight As Integer 
 Dim intTotalFormWidth As Integer 
 Dim intHeightHeader As Integer 
 Dim intHeightDetail As Integer 
 Dim intHeightFooter As Integer 
 
 ' Determine form's height. 
 intHeightHeader = frm.Section(acHeader).Height 
 intHeightDetail = frm.Section(acDetail).Height 
 intHeightFooter = frm.Section(acFooter).Height 
 intTotalFormHeight = intHeightHeader _ 
 + intHeightDetail + intHeightFooter 
 ' Determine form's width. 
 intTotalFormWidth = frm.Width 
 ' Determine window's height and width. 
 intWindowHeight = frm.InsideHeight 
 intWindowWidth = frm.InsideWidth 
 
 If intWindowWidth <> intTotalFormWidth Then 
 frm.InsideWidth = intTotalFormWidth 
 End If 
 If intWindowHeight <> intTotalFormHeight Then 
 frm.InsideHeight = intTotalFormHeight 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

