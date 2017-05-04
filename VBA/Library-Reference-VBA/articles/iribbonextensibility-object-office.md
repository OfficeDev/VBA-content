---
title: IRibbonExtensibility Object (Office)
keywords: vbaof11.chm289000
f1_keywords:
- vbaof11.chm289000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.IRibbonExtensibility
ms.assetid: b27a7576-b6f5-031e-e307-78ef5f8507e0
---


# IRibbonExtensibility Object (Office)

The interface through which the Ribbon user interface (UI) communicates with a COM add-in to customize the UI.


## Remarks

The  **IRibbonExtensibility** interface has a single method, **GetCustomUI**.


## Example

In the following example, written in C#, the  **IRibbonExtensibility** interface is specified in the class definition. The procedure then implements the interfaces's only method, **GetCustomUI**. This method creates an instance of a **SteamReader** object that reads in the customized markup stored in an external XML file.


```
public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility 
... 
public string GetCustomUI(string RibbonID) 
{ 
 StreamReader customUIReader = new System.IO.StreamReader("C:\\RibbonXSampleCS\\customUI.xml"); 
 string customUIData = customUIReader.ReadToEnd(); 
 return customUIData; 
} 

```


## Methods



|**Name**|
|:-----|
|[GetCustomUI](http://msdn.microsoft.com/library/a0106415-999e-94da-379c-70fb7aa6119f%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[IRibbonExtensibility Object Members](http://msdn.microsoft.com/library/8d8ecf4f-5502-1876-46af-381078c7710e%28Office.15%29.aspx)
