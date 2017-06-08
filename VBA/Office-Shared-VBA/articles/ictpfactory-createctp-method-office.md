---
title: ICTPFactory.CreateCTP Method (Office)
keywords: vbaof11.chm304001
f1_keywords:
- vbaof11.chm304001
ms.prod: office
api_name:
- Office.ICTPFactory.CreateCTP
ms.assetid: 17be1aa2-5045-2c89-151b-6f00d1bae6c1
ms.date: 06/08/2017
---


# ICTPFactory.CreateCTP Method (Office)

Creates an instance of a custom task pane.


## Syntax

 _expression_. **CreateCTP**( **_CTPAxID_**, **_CTPTitle_**, **_CTPParentWindow_** )

 _expression_ An expression that returns a **ICTPFactory** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CTPAxID_|Required|**String**|The CLSID or ProgID of a Microsoft ActiveX® object. |
| _CTPTitle_|Required|**String**|The title for the task pane.|
| _CTPParentWindow_|Optional|**Variant**|The window that hosts the task pane. If not present, the parent of the task pane is the ActiveWindow of the host application.|

### Return Value

CustomTaskPane


## Example

The following example, written in C#, creates an instance of a  **CustomTaskPane** object through the **ICustomTaskPaneConsumer** interface and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes a **CTPFactory** object to the add-in, which you can use during the add-in's lifetime to create task panes by using the **CreateCTP** method. Note that the example assumes that the task pane is part of an COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references an ActiveX control, SampleActiveX.myControl, that was created in a separate project.


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
```


 **Note**  You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files. For example, Microsoft Visual Basic® 6.0, Microsoft Visual Basic .NET, Microsoft Visual C++®, Microsoft Visual C++ .NET, and Microsoft Visual C#®. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also


#### Concepts


[ICTPFactory Object](ictpfactory-object-office.md)
#### Other resources


[ICTPFactory Object Members](ictpfactory-members-office.md)

