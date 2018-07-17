---
title: Window.Caption Property (Publisher)
keywords: vbapb10.chm262146
f1_keywords:
- vbapb10.chm262146
ms.prod: publisher
api_name:
- Publisher.Window.Caption
ms.assetid: 1dbf66c9-e964-b17f-684f-70cbbaa5fbc7
ms.date: 06/08/2017
---


# Window.Caption Property (Publisher)

Returns or sets a  **String** indicating the caption at the top of the Microsoft Publisher application window. Read/write.


## Syntax

 _expression_. **Caption**

 _expression_A variable that represents a  **Window** object.


### Return Value

String


## Example

The following example demonstrates how a subroutine could temporarily change the Publisher window caption and then restore it afterward.


```vb
Sub WindowCaption() 
 Dim strCaption As String 
 
 strCaption = ActiveWindow.Caption 
 
 ActiveWindow.Caption = "Custom process--please wait..." 
 
 ' Run custom code here. 
 
 ActiveWindow.Caption = strCaption 
End Sub
```


