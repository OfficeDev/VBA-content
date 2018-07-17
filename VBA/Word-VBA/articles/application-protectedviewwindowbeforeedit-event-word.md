---
title: Application.ProtectedViewWindowBeforeEdit Event (Word)
keywords: vbawd10.chm4000031
f1_keywords:
- vbawd10.chm4000031
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowBeforeEdit
ms.assetid: 1ea33944-1b2f-f914-f04a-81751cc750f8
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowBeforeEdit Event (Word)

Occurs immediately before editing is enabled on the document in the specified protected view window.


## Syntax

 _expression_ . **ProtectedViewWindowBeforeEdit**( **_PvWindow_** , **_Cancel_** )

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](protectedviewwindow-object-word.md)**|The protected view window that contains the document that is enabled for editing.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , editing is not enabled on the document.|

## Example

The following code example prompts the user for a yes or no response before enabling editing on a document in a protected view window. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowBeforeEdit(ByVal PvWindow As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 &; "want to edit the document?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

