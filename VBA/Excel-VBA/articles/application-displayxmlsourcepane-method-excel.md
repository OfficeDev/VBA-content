---
title: Application.DisplayXMLSourcePane Method (Excel)
keywords: vbaxl10.chm133294
f1_keywords:
- vbaxl10.chm133294
ms.prod: excel
api_name:
- Excel.Application.DisplayXMLSourcePane
ms.assetid: 1dea98ac-8d36-4745-cb6a-9a607e863ff2
ms.date: 06/08/2017
---


# Application.DisplayXMLSourcePane Method (Excel)

Opens the  **XML Source** task pane and displays the XML map specified by the _XmlMap_ argument.


## Syntax

 _expression_ . **DisplayXMLSourcePane**( **_XmlMap_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XmlMap_|Optional| **Variant**|The XML map to display in the task pane.|

## Remarks

You can use the following code to hide the  **XML Source** task pane.


```vb
Application.CommandBars("XML Source").Visible = False
```


## Example

The following example adds an XML map named Customers to the active workbook, and then displays the XML map in the  **XML Source** task pane.


```vb
Sub DisplayXMLMap() 
 Dim objCustomer As XmlMap 
 
 Set objCustomer = ActiveWorkbook.XmlMaps.Add( _ 
 "Customers.xsd", "Root") 
 
 objCustomer.Name = "Customers" 
 
 Application.DisplayXMLSourcePane 
 objCustomer 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

