---
title: Page.Background Property (Visio)
keywords: vis_sdr.chm10913110
f1_keywords:
- vis_sdr.chm10913110
ms.prod: visio
api_name:
- Visio.Page.Background
ms.assetid: fee785fd-2872-a64e-a80e-46034255b414
ms.date: 06/08/2017
---


# Page.Background Property (Visio)

Determines whether a page is a background page. Read/write.


## Syntax

 _expression_ . **Background**

 _expression_ A variable that represents a **Page** object.


### Return Value

Integer


## Remarks

The  **Background** property must always be true for markup pages.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to iterate through a document's pages and determine whether a page is a foreground or background page. It displays the foreground pages in a list box. To run this macro, first insert a user form containing a list box control into your project.


```vb
 
Public Sub Background_Example() 
 
 Dim vsoPages As Visio.Pages 
 Dim vsoPage As Visio.Page 
 Dim intCounter As Integer 
 
 'Get the Pages collection. 
 Set vsoPages = ThisDocument.Pages 
 
 'Make sure the list box is cleared. 
 UserForm1.ListBox1.Clear 
 
 'Iterate through the collection. 
 For intCounter = 1 To vsoPages.Count 
 
 'Retrieve the Page object at the current index. 
 Set vsoPage = vsoPages(intCounter) 
 
 'Check whether the current page is a background page. 
 'Display the names of all the foreground pages. 
 If vsoPage.Background = False Then 
 
 UserForm1.ListBox1.AddItem vsoPage.Name 
 
 End If 
 
 Next intCounter 
 
 'Display the user form. 
 UserForm1.Show 
 
End Sub
```


