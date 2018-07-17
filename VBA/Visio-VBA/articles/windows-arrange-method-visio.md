---
title: Windows.Arrange Method (Visio)
keywords: vis_sdr.chm11716080
f1_keywords:
- vis_sdr.chm11716080
ms.prod: visio
api_name:
- Visio.Windows.Arrange
ms.assetid: 0a7f5b76-d2e9-7640-f2e7-ed68ef170f77
ms.date: 06/08/2017
---


# Windows.Arrange Method (Visio)

Arranges the windows in a  **Windows** collection.


## Syntax

 _expression_ . **Arrange**( **_nArrangeFlags_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nArrangeFlags_|Optional| **Variant**|A flag that specifies how to arrange the windows; by default, the windows are arranged vertically.|

### Return Value

Nothing


## Remarks

Using the  **Arrange** method is equivalent to clicking **Arrange All** in the **Window** group on the **View** tab. The active window remains active.

Visio considers windows from top to bottom and then from left to right. You can influence which windows will end up topmost when tiling horizontally (or leftmost when tiling vertically) by prearranging windows.

The following constants declared by the Visio type library are valid values for  _nArrangeFlags_. These constants are also declared by the Visio type library in  **VisWindowArrange** .



|**Constant**|**Value**|
|:-----|:-----|
| **VisArrangeTileVertical**|1|
| **VisArrangeTileHorizontal**|2|
| **VisArrangeCascade**|3|

## Example

The following macro shows how to activate and arrange windows.


```vb
 
Public Sub Arrange_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoWindow As Visio.Window 
 Dim vsoWindow2 As Visio.Window 
 
 'Create two new windows by adding documents. 
 Set vsoDocument = Documents.Add("") 
 Set vsoWindow = ActiveWindow 
 Set vsoDocument = Documents.Add("") 
 
 'Use the Arrange method to tile the windows 
 '(currently the last opened window is active). 
 Windows.Arrange 
 
 'Use the Activate method to make the other 
 'window the active window. 
 vsoWindow.Activate 
 
End Sub
```


