---
title: Application.DocSize Method (Project)
keywords: vbapj.chm2017
f1_keywords:
- vbapj.chm2017
ms.prod: project-server
api_name:
- Project.Application.DocSize
ms.assetid: 03eb42ef-748e-ef42-a453-8305b0e2835c
ms.date: 06/08/2017
---


# Application.DocSize Method (Project)

Sets the width and height of the active window.


## Syntax

 _expression_. **DocSize**( ** _Width_**, ** _Height_**, ** _Points_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Width_|Optional|**Long**|A number that specifies the new width of the active window.|
| _Height_|Optional|**Long**|A number that specifies the new height of the active window.|
| _Points_|Optional|**Boolean**|**True** if **Width** and **Height** are measured in points. **False** if they are measured in pixels. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example tiles the windows of open projects vertically within the main window of Project.


```vb
Sub TileProjectWindowsVertically() 
 
    Dim I As Long   ' Index used in For...Next loop 
     
    For I = 1 To Application.Windows.Count 
        Windows(I).Activate 
        DocSize Width:=UsableWidth / Windows.Count, Height:=UsableHeight, Points:=True 
        DocMove XPosition:=(I - 1) * UsableWidth / Windows.Count, YPosition:=0, Points:=True 
    Next I 
End Sub
```


