---
title: CustomTaskPane Object (Office)
keywords: vbaof11.chm3030000
f1_keywords:
- vbaof11.chm3030000
ms.prod: office
api_name:
- Office.CustomTaskPane
ms.assetid: 7ed379b7-d070-4d7b-abe1-92dc73d3d137
ms.date: 06/08/2017
---


# CustomTaskPane Object (Office)

Represents a custom task pane in the container application.


## Example

The following example, written in C#, creates an instance of a  **CustomTaskPane** object and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create a task pane by using the **CreateCTP** method. Note that the example assumes that the task pane is part of an COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references a Microsoft ActiveX® control, SampleActiveX.myControl, that was created in a separate project.


```
public class Connect : Object, Extensibility.IDTExtensibility2, ICustomTaskPaneConsumer 
... 
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
} 
...
```


 **Note**  You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files. For example, Microsoft Visual Basic® 6.0, Microsoft Visual Basic .NET, Microsoft Visual C++®, Microsoft Visual C++ .NET, and Microsoft Visual C#®. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## Events



|**Name**|
|:-----|
|[DockPositionStateChange](customtaskpane-dockpositionstatechange-event-office.md)|
|[VisibleStateChange](customtaskpane-visiblestatechange-event-office.md)|

## Methods



|**Name**|
|:-----|
|[Delete](customtaskpane-delete-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](customtaskpane-application-property-office.md)|
|[ContentControl](customtaskpane-contentcontrol-property-office.md)|
|[DockPosition](customtaskpane-dockposition-property-office.md)|
|[DockPositionRestrict](customtaskpane-dockpositionrestrict-property-office.md)|
|[Height](customtaskpane-height-property-office.md)|
|[Title](customtaskpane-title-property-office.md)|
|[Visible](customtaskpane-visible-property-office.md)|
|[Width](customtaskpane-width-property-office.md)|
|[Window](customtaskpane-window-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
