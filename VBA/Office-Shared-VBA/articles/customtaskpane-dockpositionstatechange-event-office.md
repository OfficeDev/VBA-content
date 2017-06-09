---
title: CustomTaskPane.DockPositionStateChange Event (Office)
keywords: vbaof11.chm302002
f1_keywords:
- vbaof11.chm302002
ms.prod: office
api_name:
- Office.CustomTaskPane.DockPositionStateChange
ms.assetid: fd22407b-4926-2de5-ec1d-aad1a13fe269
ms.date: 06/08/2017
---


# CustomTaskPane.DockPositionStateChange Event (Office)

Occurs when the user changes the docking position of the active custom task pane.


## Syntax

 _expression_. **DockPositionStateChange**( **_CustomTaskPaneInst_**, )

 _expression_ An expression that returns a **CustomTaskPane** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomTaskPaneInst_|Required|**Object**|The active custom task pane.|

## Example

The following example, written in C#, creates a custom task pane and adds a Microsoft ActiveX® button control that was created in another project. A  **DockPositionStateChange** event of type **_CustomTaskPaneEvents_DockPositionStateChangeEventHandler** is then defined. When the event is triggered, a message box is displayed telling the user that the docked task pane has been moved.


```
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
 
 CTP.DockPositionStateChange += new _CustomTaskPaneEvents_DockPositionStateChangeEventHandler(CTP_DockPositionStateChange); 
 
} 
 
private void CTP_DockPositionStateChange(object sender, string dockpositionArgs) 
{ 
 Console.WriteLine("The custom task pane was moved"); 
}
```


 **Note**  Custom task panes can be created in any language that supports COM and allows you to create dynamic-linked library (DLL) files. For example, Microsoft Visual Basic® 6.0, Microsoft Visual Basic .NET, Microsoft Visual C++®, Microsoft Visual C++ .NET, and Microsoft Visual C#®. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also


#### Concepts


[CustomTaskPane Object](customtaskpane-object-office.md)
#### Other resources


[CustomTaskPane Object Members](customtaskpane-members-office.md)

