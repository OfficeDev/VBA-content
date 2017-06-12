---
title: Application.DisplayPasteOptions Property (Excel)
keywords: vbaxl10.chm133273
f1_keywords:
- vbaxl10.chm133273
ms.prod: excel
api_name:
- Excel.Application.DisplayPasteOptions
ms.assetid: da9cc6c1-e803-411a-220d-5c9c82d94504
ms.date: 06/08/2017
---


# Application.DisplayPasteOptions Property (Excel)

 **True** if the **Paste Options** button can be displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayPasteOptions**

 _expression_ A variable that represents an **Application** object.


## Remarks

This is a Microsoft Office-wide setting. This setting affects all other Microsoft Office applications. Setting the  **DisplayPasteOptions** property to **True** turns off the **Auto Fill Options** button in Microsoft Excel. The **Auto Fill Options** button is only in Excel, but the **Paste Options** button is in all the other Microsoft Office applications.


## Example

In this example, Microsoft Excel notifies the user the status of displaying the  **Paste Options** button.


```vb
Sub CheckDisplayFeature() 
 
 ' Check if the options button can be displayed. 
 If Application.DisplayPasteOptions = True Then 
 MsgBox "The ability to display the Paste Options button is on." 
 Else 
 MsgBox "The ability to display the Paste Options button is off." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

