---
title: Application.ProtectedViewWindowBeforeEdit Event (Excel)
keywords: vbaxl10.chm504109
f1_keywords:
- vbaxl10.chm504109
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowBeforeEdit
ms.assetid: b823b4a4-5d2f-7caf-f66f-5053b58082e4
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowBeforeEdit Event (Excel)

Occurs immediately before editing is enabled on the workbook in the specified  **Protected View** window.


## Syntax

 _expression_ . **ProtectedViewWindowBeforeEdit**( **_Pvw_** , **_Cancel_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](protectedviewwindow-object-excel.md)**|The  **Protected View** window that contains the workbook that is enabled for editing.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , editing is not enabled on the workbook.|

### Return Value

Nothing


## Example

The following code example prompts the user for a yes or no response before enabling editing on a workbook in a  **Protected View** window. This code must be placed in a class module, and an instance of the class must be correctly initialized. For more information about how to use event procedures with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_ProtectedViewWindowBeforeEdit(ByVal Pvw As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 &; "want to edit the workbook?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

