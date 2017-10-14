---
title: Window.DockedStencils Method (Visio)
keywords: vis_sdr.chm11616185
f1_keywords:
- vis_sdr.chm11616185
ms.prod: visio
api_name:
- Visio.Window.DockedStencils
ms.assetid: d16865ee-a21f-75c7-55c4-8b30f1ae91e3
ms.date: 06/08/2017
---


# Window.DockedStencils Method (Visio)

Returns the names of all stencils docked in a Microsoft Visio drawing window.


## Syntax

 _expression_ . **DockedStencils**( **_NameArray()_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameArray()_|Required| **String**|Out parameter. An array that receives the names of stencils docked in a window.|

### Return Value

Nothing


## Remarks

The  **DockedStencils** method returns an array of strings?the names of the stencils shown in the docked stencil panes of a window. When the window is a drawing window, the number of docked stencil panes ( _n_ ) is equal to or greater than zero, and _n_ is zero when the window isn't a drawing window.

If the  **DockedStencils** method succeeds, _NameArray()_ returns a one-dimensional array of _n_ strings indexed from zero (0) to _n_ - 1. The _NameArray()_ paramter is an out parameter that is allocated by the **DockedStencils** method, ownership of which is passed back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. Note that the **SafeArrayDestroy** procedure has the side effect of freeing the strings referenced by the array's entries. The **DockedStencils** method fails if _NameArray()_ is null. (Microsoft Visual Basic and Visual Basic for Applications take care of destroying the array for you.)

If  _strStencilName_ is the string returned by _NameArray(StencilName)_,  **Documents.Item** ( _strStencilName_) succeeds and returns a  **Document** object representing the stencil.


## Example

The following Microsoft Visual Basic for Applications macro shows how to use the  **DockedStencils** method to get the document names of all the stencils docked in the active window. It also prints, in the **Immediate** window, the name of the active document and the lower and upper bounds of the array that holds the stencil names, and then it lists the stencil names and paths, also in the **Immediate** window.


```vb
 
Public Sub DockedStencils_Example() 
 
 Dim astrStencilNames() As String 
 ActiveWindow.DockedStencils astrStencilNames 
 
 Dim intLowerBound As Integer 
 Dim intUpperBound As Integer 
 Dim intIndex As Integer 
 
 intLowerBound = LBound(astrStencilNames) 
 intUpperBound = UBound(astrStencilNames) 
 Debug.Print "Active document: " ActiveWindow.Document; " Lower bound:"; intLowerBound; " Upper Bound:"; intUpperBound 
 
 intIndex = intLowerBound 
 While intIndex <= intUpperBound 
 Debug.Print astrStencilNames(intIndex) 
 intIndex = intIndex + 1 
 Wend 
 
End Sub
```


