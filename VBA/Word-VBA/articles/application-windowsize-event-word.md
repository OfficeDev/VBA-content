---
title: Application.WindowSize Event (Word)
keywords: vbawd10.chm4000024
f1_keywords:
- vbawd10.chm4000024
ms.prod: word
api_name:
- Word.Application.WindowSize
ms.assetid: 96d55786-52c8-68a9-b9e9-b29c320a435a
ms.date: 06/08/2017
---


# Application.WindowSize Event (Word)

Occurs when the application window is resized or moved.


## Syntax

 _expression_ . **Private Sub object_WindowSize**( **_ByVal Doc As Document_** , **_ByVal Wn As Window_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document in the window being sized.|
| _Wn_|Required| **Window**|The window being sized.|

## Example

This example displays a message every time the Microsoft Word application window is moved or resized. This example assumes that you have declared an application variable called "WordApp" in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub WordApp_WindowSize(ByVal Doc As Document, _ 
 ByVal Wn As Window) 
 MsgBox "You have just resized or moved your window." 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

