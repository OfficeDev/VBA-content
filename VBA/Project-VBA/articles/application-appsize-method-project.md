---
title: Application.AppSize Method (Project)
keywords: vbapj.chm2012
f1_keywords:
- vbapj.chm2012
ms.prod: project-server
api_name:
- Project.Application.AppSize
ms.assetid: 31183106-d66d-235d-608c-02d3844c0e1b
ms.date: 06/08/2017
---


# Application.AppSize Method (Project)

Sets the width and height of the main window.


## Syntax

 _expression_. **AppSize**( ** _Width_**, ** _Height_**, ** _Points_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Width_|Optional|**Long**|A number that specifies the new width of the main window.|
| _Height_|Optional|**Long**|A number that specifies the new height of the main window.|
| _Points_|Optional|**Boolean**|**True** if **Width** and **Height** are measured in points. **False** if they are measured in pixels. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example moves the main window of Project to the left half of the screen.


```vb
Sub MoveMainWindowToLeftHalf() 
 
    Dim WindowHeight As Long 
     
    ' Remember the height when maximized. 
    Application.WindowState = pjMaximized 
    WindowHeight = Application.Height 
     
    AppSize Width:=UsableWidth / 2, Height:=UsableHeight, Points:=True 
    Application.Left = 0 
    ' Be sure the window uses all the available height. 
    If Application.Height < WindowHeight Then Application.Height = WindowHeight 
     
End Sub
```


