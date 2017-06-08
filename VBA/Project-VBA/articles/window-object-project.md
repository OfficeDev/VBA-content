---
title: Window Object (Project)
keywords: vbapj.chm131356
f1_keywords:
- vbapj.chm131356
ms.prod: project-server
api_name:
- Project.Window
ms.assetid: b5dcb82d-1f5a-1334-0f03-3e23d3b9d940
ms.date: 06/08/2017
---


# Window Object (Project)

Represents a window in the application or project. The  **Window** object is a member of the **[Windows](windows-object-project.md)** collection.
 


## Remarks


 **Note**  The  **Windows** collection is maintained for backward compatibility. We recommended that you use the **[Windows2](windows2-object-project.md)** collection for all new development.
 

The  **Application.Windows** collection contains all the windows in the application, whereas the **Project.Windows** collection contains only the windows in the specified project.
 

 

## Examples

 **Using the Window object**
 

 
Use  **Windows** (*Index* ), where*Index* is the window index number or window caption, to return a single **Window** object. The following example maximizes the first window in the window list.
 

 



```
Application.Windows(1).WindowState = pjMaximized
```

The window caption is the text shown in the title bar at the top of the window when the window is not maximized. The caption is also shown in the list of open files on the bottom of the  **Windows** menu. Use the **[Caption](window-caption-property-project.md)** property to set or return the window caption. Changing the window caption does not change the name of the project. The following example hides the window that contains the caption "Project1".
 

 



```
If Application.Windows(1).Caption = "Project1" Then
    Application.Windows(1).Visible = False
End If
```

 **Using the Windows collection**
 

 
Use the  **[Windows](application-windows-property-project.md)** property to return a **Windows** collection. The following example cascades all the windows that are currently displayed in Project.
 

 



```
With Application.Windows
    For I = 1 To .Count
        .Item(I).Activate
        .Item(I).Top = (I - 1) * 15
        .Item(I).Left = (I - 1) * 15
    Next I
End With
```

Use the  **[WindowNewWindow](application-windownewwindow-method-project.md)** method to create a new window and add it to the collection. The following example creates a new window for the active project.
 

 



```
Application.WindowNewWindow
```


## Methods



|**Name**|
|:-----|
|[Activate](window-activate-method-project.md)|
|[Close](window-close-method-project.md)|
|[Refresh](window-refresh-method-project.md)|
|[WebBrowserControlFrame](window-webbrowsercontrolframe-method-project.md)|
|[WebBrowserControlWindow](window-webbrowsercontrolwindow-method-project.md)|

## Properties



|**Name**|
|:-----|
|[ActivePane](window-activepane-property-project.md)|
|[Application](window-application-property-project.md)|
|[BottomPane](window-bottompane-property-project.md)|
|[Caption](window-caption-property-project.md)|
|[Height](window-height-property-project.md)|
|[Index](window-index-property-project.md)|
|[Left](window-left-property-project.md)|
|[Parent](window-parent-property-project.md)|
|[Top](window-top-property-project.md)|
|[TopPane](window-toppane-property-project.md)|
|[Visible](window-visible-property-project.md)|
|[Width](window-width-property-project.md)|
|[WindowState](window-windowstate-property-project.md)|

