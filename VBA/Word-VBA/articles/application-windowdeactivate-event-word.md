---
title: Application.WindowDeactivate Event (Word)
keywords: vbawd10.chm4000010
f1_keywords:
- vbawd10.chm4000010
ms.prod: word
api_name:
- Word.Application.WindowDeactivate
ms.assetid: 70b86ecc-40ba-6e70-b430-4fce6083ff2d
ms.date: 06/08/2017
---


# Application.WindowDeactivate Event (Word)

Occurs when any document window is deactivated.


## Syntax

 _expression_ . **Private Sub object_WindowDeactivate**( **_ByVal Doc As Document_** , **_ByVal Wn As Window_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For more information about using events with the **Application** object or the **Document** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx) or[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document displayed in the deactivated window.|
| _Wn_|Required| **Window**|The deactivated window.|

## Example

This example minimizes any document window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowDeactivate _ 
 (ByVal Wn As Word.Window) 
 Wn.WindowState = wdWindowStateMinimize 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

