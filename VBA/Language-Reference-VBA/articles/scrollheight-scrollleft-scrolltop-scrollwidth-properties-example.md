---
title: ScrollHeight, ScrollLeft, ScrollTop, ScrollWidth Properties Example
keywords: fm20.chm5225138
f1_keywords:
- fm20.chm5225138
ms.prod: office
ms.assetid: 79f36650-9779-1ae4-678c-9f239e1306e1
ms.date: 06/08/2017
---


# ScrollHeight, ScrollLeft, ScrollTop, ScrollWidth Properties Example

The following example uses a page of a  **MultiPage** as a scrolling region. The user can use the scroll bars on Page2 of the **MultiPage** to gain access to parts of the page that are not initially displayed.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **MultiPage** named MultiPage1, and that each page of the **MultiPage** contains one or more controls.

 **Note**  Each page of a  **MultiPage** is unique. Page1 has no scroll bars. Page2 has horizontal and vertical scroll bars.




```vb
Private Sub UserForm_Initialize() 
 MultiPage1.Pages(1).ScrollBars = fmScrollBarsBoth 
 MultiPage1.Pages(1).KeepScrollBarsVisible = _ 
 fmScrollBarsNone 
 
 MultiPage1.Pages(1).ScrollHeight = 2 * _ 
 MultiPage1.Height 
 MultiPage1.Pages(1).ScrollWidth = 2 * _ 
 MultiPage1.Width 
 
 'Set ScrollHeight, ScrollWidth before setting 
 'ScrollLeft, ScrollTop 
 MultiPage1.Pages(1).ScrollLeft = _ 
 MultiPage1.Width / 2 
 MultiPage1.Pages(1).ScrollTop = _ 
 MultiPage1.Height / 2 
End Sub
```


