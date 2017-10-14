---
title: HeaderFooter.IsHeader Property (Publisher)
keywords: vbapb10.chm7471109
f1_keywords:
- vbapb10.chm7471109
ms.prod: publisher
api_name:
- Publisher.HeaderFooter.IsHeader
ms.assetid: b652fcc8-2c89-6d4f-c366-4c78681bea59
ms.date: 06/08/2017
---


# HeaderFooter.IsHeader Property (Publisher)

 **True** if the specified **HeaderFooter** object is a header, **False** if it is a footer. Read-only **Boolean**.


## Syntax

 _expression_. **IsHeader**

 _expression_A variable that represents an  **HeaderFooter** object.


### Return Value

Boolean


## Example

The following example creates a new collection and fills it with the headers and footer from each master page. The collection is then iterated and a test is made to determine whether the  **HeaderFooter** object is a header or a footer, then appropriate text is written to the header or footer.


```vb
Dim objHeadersFooters As Collection 
Dim objMasterPage As page 
Dim objHeadFoot As HeaderFooter 
 
Set objHeadersFooters = New Collection 
For Each objMasterPage In ActiveDocument.masterPages 
 objHeadersFooters.Add objMasterPage.Header 
 objHeadersFooters.Add objMasterPage.Footer 
Next objMasterPage 
 
For Each objHeadFoot In objHeadersFooters 
 If objHeadFoot.IsHeader = True Then 
 objHeadFoot.TextRange.Text = "Header Text" 
 ElseIf objHeadFoot.IsHeader = False Then 
 objHeadFoot.TextRange.Text = "Footer Text" 
 End If 
Next 

```


