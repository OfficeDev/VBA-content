---
title: Application.AutoRecover Property (Excel)
keywords: vbaxl10.chm133276
f1_keywords:
- vbaxl10.chm133276
ms.prod: excel
api_name:
- Excel.Application.AutoRecover
ms.assetid: bc2453fa-4319-c1da-5ad5-2efb306c3063
ms.date: 06/08/2017
---


# Application.AutoRecover Property (Excel)

Returns an  **[AutoRecover](autorecover-object-excel.md)** object, which backs up all file formats on a timed interval.


## Syntax

 _expression_ . **AutoRecover**

 _expression_ A variable that represents an **Application** object.


## Remarks

Valid time intervals are whole numbers from 1 to 120.


## Example

In this example, the  **[Time](autorecover-time-property-excel.md)** property is used in conjunction with the **AutoRecover** property to set the time interval for Microsoft Excel to wait before saving another copy to five minutes.


```vb
Sub UseAutoRecover() 
 
 Application.AutoRecover.Time = 5 
 
 MsgBox "The time that will elapse between each automatic " &; _ 
 "save has been set to " &; _ 
 Application.AutoRecover.Time &; " minutes." 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

