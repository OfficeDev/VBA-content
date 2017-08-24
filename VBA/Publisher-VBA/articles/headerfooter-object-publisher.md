---
title: HeaderFooter Object (Publisher)
keywords: vbapb10.chm7536639
f1_keywords:
- vbapb10.chm7536639
ms.prod: publisher
api_name:
- Publisher.HeaderFooter
ms.assetid: d38e5e7e-45d7-667b-b6f2-9ad8e764af79
ms.date: 06/08/2017
---


# HeaderFooter Object (Publisher)

Represents the header or footer of a master page.
 


## Example

Use  **MasterPages.Header** or **MasterPages.Footer** to return a **HeaderFooter** object. The following example adds text to the header of the first master page of the active document.
 

 

```
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
objHeader.TextRange.Text = "Master Page 1 Header" 

```

Use  **HeaderFooter.Delete** to delete any existing content from a header or footer. Calling this method does not delete the text frame, just the contents of it. The following example deletes all of the header and footer content of all the master pages in a publication.
 

 



```
Dim objMasterPage As page 
For Each objMasterPage In ActiveDocument.masterPages 
 objMasterPage.Header.Delete 
 objMasterPage.Footer.Delete 
Next
```

Use  **HeaderFooter.TextRange** to return a **TextRange** object representing the header or footer of a master page. Any header or footer content manipulation is done with through this property of the **HeaderFooter** object. The following example first deletes any existing content and then adds some boilerplate text to the header of a master page.
 

 



```
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
With objHeader 
 .Delete 
 .TextRange.Text = "<Insert Address Here>" 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](headerfooter-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](headerfooter-application-property-publisher.md)|
|[IsHeader](headerfooter-isheader-property-publisher.md)|
|[Parent](headerfooter-parent-property-publisher.md)|
|[TextRange](headerfooter-textrange-property-publisher.md)|

