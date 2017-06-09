---
title: ICustomTaskPaneConsumer Object (Office)
keywords: vbaof11.chm305000
f1_keywords:
- vbaof11.chm305000
ms.prod: office
api_name:
- Office.ICustomTaskPaneConsumer
ms.assetid: 54be3f78-4e5d-8595-d369-0724df0debf7
ms.date: 06/08/2017
---


# ICustomTaskPaneConsumer Object (Office)

An interface that provides access to the  **CTPFactoryAvailable** method that is used to create an instance of a custom task pane.


## Example

The following example, written in C#, creates an instance of a  **CustomTaskPane** object through the **ICustomTaskPaneConsumer** interface and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **CTPFactory** object to the add-in, which you can use during the add-in's lifetime to create task panes by using the **CreateCTP** method. Note that the example assumes that the task pane is part of an COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references a Microsoft ActiveX® control, SampleActiveX.myControl, that is created in a separate project.


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
|[CTPFactoryAvailable](icustomtaskpaneconsumer-ctpfactoryavailable-method-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
