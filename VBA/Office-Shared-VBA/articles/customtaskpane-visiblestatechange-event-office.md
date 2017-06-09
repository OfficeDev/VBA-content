---
title: CustomTaskPane.VisibleStateChange Event (Office)
keywords: vbaof11.chm302001
f1_keywords:
- vbaof11.chm302001
ms.prod: office
api_name:
- Office.CustomTaskPane.VisibleStateChange
ms.assetid: 6faccef7-f35f-d0c8-383f-54493e4b4c8b
ms.date: 06/08/2017
---


# CustomTaskPane.VisibleStateChange Event (Office)

Occurs when the user changes the visibility of the custom task pane.


## Syntax

 _expression_. **VisibleStateChange**( **_CustomTaskPaneInst_**, )

 _expression_ An expression that returns a **CustomTaskPane** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomTaskPaneInst_|Required|**CustomTaskPane**|The active task pane.|

## Example

The following example, written in C#, creates a custom task pane and adds an ActiveX button control created in another project. A  **VisibleStateChange** event of type **_CustomTaskPaneEvents_VisibleStateChangeEventHandler** is defined in the procedure. When the event is triggered, the event handler displays a message box depending on whether the task pane is currently visible or hidden.


```
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
 
 CTP.VisibleStateChange += new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(CTP_VisibleStateChange); 
} 
 
private void CTP_VisibleStateChange(object sender, string visiblestateArgs) 
{ 
 if (CTP.Visible) 
 { 
 Console.WriteLine("The custom task pane is now visible"); 
 } 
 else 
 { 
 Console.WriteLine("The custom task pane has been hidden"); 
 } 
} 

```


 **Note**  Custom task panes can be created in any language that supports COM and allows you to create dynamic-linked library (DLL) files. For example, Microsoft Visual Basic® 6.0, Microsoft Visual Basic .NET, Microsoft Visual C++®, Microsoft Visual C++ .NET, and Microsoft Visual C#®. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also


#### Concepts


[CustomTaskPane Object](customtaskpane-object-office.md)
#### Other resources


[CustomTaskPane Object Members](customtaskpane-members-office.md)

