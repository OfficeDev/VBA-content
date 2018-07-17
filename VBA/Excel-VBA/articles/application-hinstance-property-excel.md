---
title: Application.Hinstance Property (Excel)
keywords: vbaxl10.chm133278
f1_keywords:
- vbaxl10.chm133278
ms.prod: excel
api_name:
- Excel.Application.Hinstance
ms.assetid: 4551a0a2-0730-1288-7a13-b2beff2a2fca
ms.date: 06/08/2017
---


# Application.Hinstance Property (Excel)

Returns a handle to the instance of Excel represented by the [Application](application-object-excel.md) object. Read-only **Long** .


## Syntax

 _expression_ . **Hinstance**

 _expression_ A variable that represents an **Application** object.


## Remarks


 **Important**  This property returns a correct handle only in the 32-bit version of Excel. In Excel, the [HinstancePtr](application-hinstanceptr-property-excel.md) property was introduced, which works correctly in both 32- and 64-bit versions of Excel.


## Example

In this example, a message box displays the Microsoft Excel instance handle to the user.


```vb
Sub CheckHinstance() 
 
 MsgBox Application.Hinstance 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

