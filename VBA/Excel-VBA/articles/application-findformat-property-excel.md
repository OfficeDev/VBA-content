---
title: Application.FindFormat Property (Excel)
keywords: vbaxl10.chm133262
f1_keywords:
- vbaxl10.chm133262
ms.prod: excel
api_name:
- Excel.Application.FindFormat
ms.assetid: b2b62232-1f11-ec82-9344-edd39e0ae33d
ms.date: 06/08/2017
---


# Application.FindFormat Property (Excel)

Sets or returns the search criteria for the type of cell formats to find.


## Syntax

 _expression_ . **FindFormat**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, the search criteria is set to look for Arial, Regular, Size 10 font cells and the user is notified.


```vb
Sub UseFindFormat() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Notify user. 
 With Application.FindFormat.Font 
 MsgBox .Name &; "-" &; .FontStyle &; "-" &; .Size &; _ 
 " font is what the search criteria is set to." 
 End With 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

