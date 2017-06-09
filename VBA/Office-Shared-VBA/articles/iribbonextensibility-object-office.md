---
title: IRibbonExtensibility Object (Office)
keywords: vbaof11.chm289000
f1_keywords:
- vbaof11.chm289000
ms.prod: office
api_name:
- Office.IRibbonExtensibility
ms.assetid: b27a7576-b6f5-031e-e307-78ef5f8507e0
ms.date: 06/08/2017
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


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[IRibbonExtensibility Object Members](iribbonextensibility-members-office.md)

