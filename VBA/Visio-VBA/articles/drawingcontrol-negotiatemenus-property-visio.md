---
title: DrawingControl.NegotiateMenus Property (Visio)
keywords: vis_sdr.chm52005
f1_keywords:
- vis_sdr.chm52005
ms.prod: visio
api_name:
- Visio.NegotiateMenus
ms.assetid: 03a9ee41-1116-101f-fdbe-16b8fcd98866
ms.date: 06/08/2017
---


# DrawingControl.NegotiateMenus Property (Visio)

Determines whether Microsoft Visio menus are merged with those of the host application when the Microsoft Visio Drawing Control has the keyboard focus. Read/write.


## Syntax

 _expression_ . **NegotiateMenus**

 _expression_ A variable that represents a **DrawingControl** object.


### Return Value

Boolean


## Remarks

You can set the  **NegotiateMenus** property either at design time (for example, in the **Properties** window in Microsoft Visual Basic 6.0), or at run time, as shown in the following example. It is recommended that you set **NegotiateMenus** at design time. If you do set **NegotiateMenus** at run time, set the property prior to in-place activation of the Visio Drawing Control.

When  **NegotiateMenus** is set to **True** , you can use the methods and properties of the Visio object model to customize Visio menus in the Visio Drawing Control window.

Visio task panes are implemented as toolbars. If you set  **NegotiateMenus** to **True** , but set **NegotiateToolbars** to **False** , menu commands such as **Task Pane** ( **View** menu) and **Microsoft Visio Help** ( **Help** menu) will be unavailable.




 **Note**  If  **NegotiateMenus** is **True** , the Visio Drawing Control supports OLE2 menu merging. For this merging to function properly, the host container application must implement the **IOleInPlaceFrame** interface correctly. For more information, search for "IOleInPlaceFrame" on MSDN.

 When there is only a single instance of the control in your application, if you set the **NegotiateMenus** property to **True** and the **NegotiateToolbars** property to **False** , or vice versa, Visio task panes will not be displayed as expected. In order for Visio task panes to appear in the Visio Drawing Control, both properties must be set to the same value.

However, if your application uses multiple instances of the Visio Drawing Control, you can set either the  **NegotiateMenus** property or the **NegotiateToolbars** property to **True** , but not both. If both properties are set to **True** when you are using multiple instances of the control in a single application, you will experience unexpected behavior.


## Example

The following example shows how to set the  **NegotiateMenus** property at run time in the **Form_Load()** sub procedure of your Visual Basic program. For examples of how to display or modify one or more particular Visio menus, see the topics for the **UIObject** object and its member methods and properties, in this Automation Reference.


```vb
Private Sub Form_Load() 
 
 vsoDrawingControl.NegotiateMenus = True 
 
End Sub
```


