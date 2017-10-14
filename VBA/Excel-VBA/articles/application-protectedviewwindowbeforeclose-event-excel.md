---
title: Application.ProtectedViewWindowBeforeClose Event (Excel)
keywords: vbaxl10.chm504110
f1_keywords:
- vbaxl10.chm504110
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowBeforeClose
ms.assetid: 5fa37062-61c7-3002-1ea0-c5bd396b6a9b
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowBeforeClose Event (Excel)

Occurs immediately before a  **Protected View** window or a workbook in a **Protected View** window closes.


## Syntax

 _expression_ . **ProtectedViewWindowBeforeClose**( **_Pvw_** , **_Reason_** , **_Cancel_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](protectedviewwindow-object-excel.md)**|An object that represents the  **Protected View** window that is closed.|
| _Reason_|Required| **[XlProtectedViewCloseReason](xlprotectedviewclosereason-enumeration-excel.md)**|A constant that specifies the reason the  **Protected View** window is closed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the window does not close when the procedure is finished.|

### Return Value

Nothing


## Example

The following code example prompts the user for a yes or no response before closing the  **Protected View** window. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_ProtectedViewWindowBeforeClose(ByVal Pvw as ProtectedViewWindow, _ 
 Reason as XlProtectedViewCloseReason, Cancel as Boolean) 
 a = MsgBox("Do you really want to close the Protected View window?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

