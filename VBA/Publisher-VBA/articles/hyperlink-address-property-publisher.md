---
title: Hyperlink.Address Property (Publisher)
keywords: vbapb10.chm4587523
f1_keywords:
- vbapb10.chm4587523
ms.prod: publisher
api_name:
- Publisher.Hyperlink.Address
ms.assetid: 784a9213-38bc-c5fd-f215-abeb174ec628
ms.date: 06/08/2017
---


# Hyperlink.Address Property (Publisher)

Returns or sets a  **String** that represents the URL address for a hyperlink. Read/write.


## Syntax

 _expression_. **Address**

 _expression_A variable that represents a  **Hyperlink** object.


### Return Value

String


## Example

This example displays the URL addresses for all hyperlinks in the active publication.


```vb
Sub ShowHyperlinkAddresses() 
 Dim pgsPage As Page 
 Dim shpShape As Shape 
 Dim hprLink As Hyperlink 
 Dim intCount As Integer 
 For Each pgsPage In ActiveDocument.Pages 
 For Each shpShape In pgsPage.Shapes 
 If shpShape.TextFrame.TextRange.Hyperlinks.Count > 0 Then 
 For Each hprLink In shpShape.TextFrame.TextRange.Hyperlinks 
 MsgBox "This hyperlink goes to " &; hprLink.Address &; "." 
 intCount = intCount + 1 
 Next hprLink 
 ElseIf shpShape.Hyperlink.Address <> "" Then 
 MsgBox "This hyperlink goes to " &; shpShape.Hyperlink.Address &; "." 
 intCount = intCount + 1 
 End If 
 Next shpShape 
 Next pgsPage 
 If intCount < 1 Then 
 MsgBox "You don't have any hyperlinks in your publication." 
 Else 
 MsgBox "You have " &; intCount &; " hyperlinks in " &; ThisDocument.Name &; "." 
 End If 
End Sub
```


