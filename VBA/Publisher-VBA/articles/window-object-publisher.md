---
title: Window Object (Publisher)
keywords: vbapb10.chm327679
f1_keywords:
- vbapb10.chm327679
ms.prod: publisher
api_name:
- Publisher.Window
ms.assetid: 342d77cd-5556-6ac3-a828-b1b60380f910
ms.date: 06/08/2017
---


# Window Object (Publisher)

Represents a window. Many publication characteristics, such as scroll bars and rulers, are actually properties of the window.
 


## Example

Use the  **[ActiveWindow](application-activewindow-property-publisher.md)** property to return a **Window** object. The following example maximizes the active window.
 

 

```
Sub MaximizeWindow 
 ActiveWindow.WindowState = pbWindowStateMaximize 
End Sub
```

Use the  **[Caption](window-caption-property-publisher.md)** property to return the file and application names of the active window. The following example displays a message with the file name and Microsoft Publisher application name.
 

 



```
Sub ShowFileApNames 
 MsgBox Windows(1).Caption 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Activate](window-activate-method-publisher.md)|
|[Move](window-move-method-publisher.md)|
|[Resize](window-resize-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](window-application-property-publisher.md)|
|[Caption](window-caption-property-publisher.md)|
|[Height](window-height-property-publisher.md)|
|[Hwnd](window-hwnd-property-publisher.md)|
|[Left](window-left-property-publisher.md)|
|[Parent](window-parent-property-publisher.md)|
|[Top](window-top-property-publisher.md)|
|[Visible](window-visible-property-publisher.md)|
|[Width](window-width-property-publisher.md)|
|[WindowState](window-windowstate-property-publisher.md)|

