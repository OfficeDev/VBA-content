---
title: Form.InsideHeight Property (Access)
keywords: vbaac10.chm13475
f1_keywords:
- vbaac10.chm13475
ms.prod: access
api_name:
- Access.Form.InsideHeight
ms.assetid: 7a49b4b4-1bbf-c0ad-d873-ff81f8b99929
ms.date: 06/08/2017
---


# Form.InsideHeight Property (Access)

You can use the  **InsideHeight** property (along with the **InsideWidth** property) to determine the height and width (in twips) of the window containing a form. Read/write **Long**.


## Syntax

 _expression_. **InsideHeight**

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

