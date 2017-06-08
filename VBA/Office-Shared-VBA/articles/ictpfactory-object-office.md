---
title: ICTPFactory Object (Office)
keywords: vbaof11.chm304000
f1_keywords:
- vbaof11.chm304000
ms.prod: office
api_name:
- Office.ICTPFactory
ms.assetid: da653cf7-9649-dc07-e3ae-4f7805fe3eb1
ms.date: 06/08/2017
---


# ICTPFactory Object (Office)

Used to create a custom task pane.


## Example

The following example, written in C#, creates an instance of a  **CustomTaskPane** object and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create a task pane by using the **CreateCTP** method. Note that the example assumes that the task pane is part of an COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references a Microsoft ActiveX® control, SampleActiveX.myControl, that is created in a separate project.


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


## Methods



|**Name**|
|:-----|
|[CreateCTP](ictpfactory-createctp-method-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
