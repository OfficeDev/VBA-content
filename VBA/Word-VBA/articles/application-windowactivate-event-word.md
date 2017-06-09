---
title: Application.WindowActivate Event (Word)
keywords: vbawd10.chm400009
f1_keywords:
- vbawd10.chm400009
ms.prod: word
api_name:
- Word.Application.WindowActivate
ms.assetid: f1340e1e-6aec-edaa-78c2-47e3e1d5299f
ms.date: 06/08/2017
---


# Application.WindowActivate Event (Word)

Occurs when any document window is activated.


## Syntax

 _expression_ . **Private Sub object_WindowActivate**( **_ByVal Doc As Document_** , **_ByVal Wn As Window_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For more information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document displayed in the activated window.|
| _Wn_|Required| **Window**|The window that's being activated.|

## Example

This example maximizes any document window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowActivate _ 
 (ByVal Doc As Word.Document, _
  ByVal Wn As Word.Window) 
 Wn.WindowState = wdWindowStateMaximize 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

