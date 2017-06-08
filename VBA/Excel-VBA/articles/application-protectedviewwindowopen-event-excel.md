---
title: Application.ProtectedViewWindowOpen Event (Excel)
keywords: vbaxl10.chm504108
f1_keywords:
- vbaxl10.chm504108
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowOpen
ms.assetid: 17c847d9-a9d2-28da-832a-01d7719f1248
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowOpen Event (Excel)

Occurs when a workbook is opened in a  **Protected View** window.


## Syntax

 _expression_ . **ProtectedViewWindowOpen**( **_Pvw_** , )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](protectedviewwindow-object-excel.md)**|An object that represents the  **Protected View** window that is opened.|

### Return Value

Nothing


## Example

The following code example informs the user that the workbook will be opened in a  **Protected View** window. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow) 
 MsgBox "You are opening the following workbook in Protected View: " _ 
 &; Pvw.Caption 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

