---
title: Application.ProtectedViewWindowDeactivate Event (Excel)
keywords: vbaxl10.chm504113
f1_keywords:
- vbaxl10.chm504113
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowDeactivate
ms.assetid: 39df50ca-53e0-784a-a803-e9ac6f456d11
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowDeactivate Event (Excel)

Occurs when a  **Protected View** window is deactivated.


## Syntax

 _expression_ . **ProtectedViewWindowDeactivate**( **_Pvw_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](protectedviewwindow-object-excel.md)**|An object that represents the deactivated  **Protected View** window.|

### Return Value

 **Nothing**


## Example

The following code example minimizes any  **Protected View** window when it is deactivated. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_ProtectedViewWindowDeactivate(ByVal Pvw As ProtectedViewWindow) 
 Pvw.WindowState = xlProtectedViewWindowMinimized 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

