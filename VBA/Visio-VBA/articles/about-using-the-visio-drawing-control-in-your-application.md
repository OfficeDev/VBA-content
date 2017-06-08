---
title: About Using the Visio Drawing Control in Your Application
keywords: vis_sdr.chm1046843
f1_keywords:
- vis_sdr.chm1046843
ms.prod: visio
ms.assetid: 04e63921-6a82-deef-f9e3-eaadcfdfdc44
ms.date: 06/08/2017
---


# About Using the Visio Drawing Control in Your Application

Visio includes an ActiveX control, the Visio Drawing Control. 

Using this control, you can embed the full functionality of the Visio drawing surface into your applications. You can take advantage of the full Visio object model (API) and you can pick the aspects of the Visio user interface you want to expose to better integrate Visio seamlessly into the user interface of your application.

 **Note**  You can embed the Visio Drawing Control in Visual Basic 6.0, Visual C++ 6.0, Visual Studio, and other ActiveX control containers. However, you cannot embed the Visio Drawing Control in another Visio drawing, another ActiveX control, a Visual Basic for Applications (VBA) form in Visio, or a Visio solution window.


## Getting started

To install the Visio Drawing Control, install Visio. When you install Visio, you can choose various installation options, including the  **Minimal Install** option. If you want to minimize the installation file size of Visio on your computer, you can choose **Minimal Install**, which installs only the minimum required Visio components, including the Visio drawing application and the Visio Drawing Control. This installation option does not include Visio solutions or Visio Help (which includes the Automation Reference and ShapeSheet Reference). 

To add the Visio Drawing Control to the  **Toolbox** in Visual Basic 6.0, on the **Project** menu, click **Components**, and then in the  **Controls** list, select **Microsoft Visio 15.0 Drawing Control Type Library**. To make the control available in other development environments, consult the documentation that comes with your development product.

Once you have opened a  **Standard EXE** project in Visual Basic and added the control to the **Toolbox**, double-click the control's icon to add an instance of the control to the form in your application. You can add multiple instances of the control, but they will share the same underlying Visio  **Application** object. As a result, programmability objects and settings associated with one instance of the control will be reflected in other instances. For example, the **Documents** collection of the **Application** object will include the **Document** object associated with each instance of the control.


## Opening a Visio drawing in the control

By default, the control opens a blank Visio document (drawing). However, you can specify, either at design time or at run time, that the control load an existing Visio document. The document you specify must be available to your users, either because you supply it along with your application, or because it exists on a network share they have access to, on an intranet site, or on the Web. To specify a document at design time, set the  **Src** property in the **Properties** window in your Visual Basic project. This is the preferred method. To specify a drawing at run time, set the **Src** property in your code, usually in the **Form_Load()** procedure. More information about using the custom properties of the Visio Drawing Control is provided later in this topic, and in the **Src** property topic in this Automation Reference.

When you set the  **SRC** property to load a file into the Visio Drawing Control, the control opens a copy of the file, but does not keep the original file open for writing. As a result, you cannot use the **Document.Save** method to save changes to a file loaded into the Visio Drawing Control. To save changes in a file, first use the **SRC** property to load the file into the Visio Drawing Control, and then set **SRC** to an empty string (""). To save the modified file to disk, use the **Document.SaveAs** method.

If you do not set the  **SRC** property to an empty string after loading a drawing into the Visio Drawing Control, each time you close and reopen your application, the original drawing will be loaded, and any modifications you or your users have made will be lost.

By default, the control does not display the Visio startup screen or the  **Available Templates** tab on startup. Furthermore, by default the control does not display a docked stencil pane on startup, but if you use the **Src** property to specify a drawing that already displays a docked stencil pane, that pane will be visible in the Visio Drawing Control window. To display the stencil pane in a blank drawing, use the **Document.OpenStencilWindow** method from the Visio object model.

By default, neither Visio menus nor Visio toolbars are displayed in the control (although shortcut menus are available). However, you can use the  **NegotiateMenus** and **NegotiateToolbars** properties of the control to display these items. More information about using the custom properties of the Visio Drawing Control is provided later in this topic, and in the **NegotiateMenus** property and **NegotiateToolbars** property topics in this Automation Reference.


 **Note**  Starting in Microsoft Visio 2010, the Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio 2013, but they function differently.

You can insert multiple instances of the Visio Drawing Control in your application. However, each instance of the control can display only one drawing window and one document. 


## Gaining access to the Visio object model

To gain access to the Visio object model, use the  **Window** or **Document** property of the Visio Drawing Control. Use the following code to get a Visio **Window** object:


```vb
Dim vsoWindow As Visio.Window
Set vsoWindow = DrawingControl1.Window
```

Use the following code to get a Visio  **Document** object:




```vb
Dim vsoDocument As Visio.Document
Set vsoDocument = DrawingControl1.Document
```

Once you have either of these objects, you can use the  **Application** property of either object to get the Visio **Application** object:




```
vsoWindow.Application
vsoDocument.Application
```

With these objects you have access to all the rest of the Visio object model. For example, if you want to modify the Visio user interface to display only the white Visio drawing surface, without the grid, scrollbars, rulers, and page tabs, you can use the following code in the  **Form_Load()** procedure:




```vb
Dim vsoWindow As Visio.Window
Set vsoWindow = DrawingControl1.Window
vsoWindow.ShowGrid = False
vsoWindow.ShowPageTabs = False
vsoWindow.ShowRulers = False
vsoWindow.ShowScrollBars = False
```

Visio events, including keyboard and mouse events, are available directly as members of the  **DrawingControl** object.

Once you have access to the Visio object model, you can use all the standard objects, methods, properties, and events of the model to automate and customize the control in your program. For more information about using the objects and members of the Visio object model, see the specific object or member topic in this Automation Reference.

Because it is intended to be embedded in another application, the Visio Drawing control does not expose the Visual Basic Editor in Visio. As a result, Visual Basic for Applications (VBA) macros in an existing Visio drawing that is opened in the control will not run.

In addition, the Visio Drawing Control does not expose the Visio ShapeSheet in the user interface. However, you can use Automation to get and set values and formulas in ShapeSheet cells.


## Using the custom properties of the Visio Drawing Control

The following table describes the custom properties exposed by the Visio Drawing Control.



|**Property**|**Description**|
|:-----|:-----|
| **Document**|Read-only. Returns the Visio  **Document** object associated with the instance of the Visio Drawing Control.|
| **HostID**|Read/write.  **String**. Returns or sets the GUID or other string assigned to the registry key that identifies the host container application (your executable program). The default is an empty string.|
| **NegotiateMenus**|Read/write.  **Boolean**. Specifies whether the control can merge menus with those of the host container application. The default is  **False**.|
| **NegotiateToolbars**|Read/write.  **Boolean**. Specifies whether the control can merge toolbars with those of the host container application. The default is  **False**.|
| **PageSizingBehavior**|Read/write. Enumerated type  **VisPageSizingBehavior**. Specifies how pages are sized and how shapes are sized and positioned when existing Visio drawings are loaded into instances of the control.|
| **Src**|Read/write.  **String**. Specifies the path to and file name of the existing Visio drawing that is loaded into an instance of the control at run time. The default is an empty string.|
| **Window**|Read-only. Returns the Visio  **Window** object associated with the instance of the Visio Drawing Control. The **Window** property is accessible only when the control is in-place active.|
For more information about any of these custom properties, and to view code examples that show how to use them, see the specific topics associated with these properties in this Automation Reference.


## Using keyboard and mouse events with the Visio Drawing Control

Beginning with Visio 2003, new keyboard and mouse events added to the Visio object model give your program the ability to respond to user keyboard and mouse input in the control. For example, you can listen for mouse clicks specific shapes in the control and write code to handle them. (For more information about how to use these events to drive actions in your host application, see the next section in this topic.)

The following new events are available:


-  **KeyDown**
    
-  **KeyPress**
    
-  **KeyUp**
    
-  **MouseDown**
    
-  **MouseMove**
    
-  **MouseUp**
    
These events are similar to the Visual Basic events that have the same names, although they take different arguments. To view the syntax, and for additional information about these events, see the specific topics associated with them in this Automation Reference. For more information about the Visual Basic events, consult Visual Basic Help.


## Integrating the Visio Drawing Control into the user interface of your application

You can use events or status changes in your host application to modify a drawing in the Visio Drawing Control. In addition, you can use events in the Visio Drawing Control to drive actions or changes in your host application. For example, you can use mouse events or keyboard events in the Visio Drawing Control to display user interface elements such as forms and message boxes in your host application. The following code shows how to handle a  **MouseDown** event (a mouse click) in the Visio Drawing Control to display a message box in your Visual Basic 6.0 application.


```vb
Private Sub DrawingControl1_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
  MsgBox "You have clicked the mouse.", , "Drawing Control Event"
 
End Sub
```


