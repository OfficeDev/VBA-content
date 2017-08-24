---
title: Application.TemplateFolderPath Property (Publisher)
keywords: vbapb10.chm131120
f1_keywords:
- vbapb10.chm131120
ms.prod: publisher
api_name:
- Publisher.Application.TemplateFolderPath
ms.assetid: e2256af9-9432-6205-864a-10bb7dec41c9
ms.date: 06/08/2017
---


# Application.TemplateFolderPath Property (Publisher)

Returns a  **String** that represents the location where Microsoft Publisher templates are stored. Read-only.


## Syntax

 _expression_. **TemplateFolderPath**

 _expression_A variable that represents a  **Application** object.


### Return Value

String


## Example

This example creates a new publication and edits the master page to contain a page number in a star in the upper-left corner of the page; then it saves the new publication to the template folder location so that it can be used as a template.


```vb
Sub CreateNewPubTemplate() 
 Dim AppPub As Application 
 Dim DocPub As Document 
 Dim strFolder As String 
 
 Set AppPub = New Publisher.Application 
 Set DocPub = AppPub.NewDocument 
 AppPub.ActiveWindow.Visible = True 
 strFolder = AppPub.TemplateFolderPath 
 
 With DocPub 
 With .MasterPages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 With .Font 
 .Bold = msoTrue 
 .Color.RGB = RGB(Red:=255, Green:=255, Blue:=255) 
 .Size = 12 
 End With 
 End With 
 End With 
 .SaveAs FileName:=strFolder &; "\NewPubTemplt.pub" 
 End With 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

