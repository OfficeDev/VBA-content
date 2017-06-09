---
title: Application.ProtectedViewWindowOpen Event (Word)
keywords: vbawd10.chm4000030
f1_keywords:
- vbawd10.chm4000030
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowOpen
ms.assetid: 42126a64-0227-d006-760e-ec11c59ef533
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowOpen Event (Word)

Occurs when a protected view window is opened.


## Syntax

 _expression_ . **ProtectedViewWindowOpen**( **_PvWindow_** , )

 _expression_ An expression that returns a **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](protectedviewwindow-object-word.md)**|The protected view window that is opened.|

## Example

The following code example informs the user that the document will be opened in a protected view window. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowOpen(ByVal PvWindow As ProtectedViewWindow) 
Dim intResponse As Integer 
 
 MsgBox "You are opening a document in " _ 
 &; "protected view window mode." 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

