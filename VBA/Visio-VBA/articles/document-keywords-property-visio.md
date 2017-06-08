---
title: Document.Keywords Property (Visio)
keywords: vis_sdr.chm10513795
f1_keywords:
- vis_sdr.chm10513795
ms.prod: visio
api_name:
- Visio.Document.Keywords
ms.assetid: c7717a93-c64f-8363-69a7-7e9ed40865dc
ms.date: 06/08/2017
---


# Document.Keywords Property (Visio)

Gets or sets the value of the  **Keywords** box in a document's **Properties** dialog box. Read/write.


## Syntax

 _expression_ . **Keywords**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Keywords** property is equivalent to entering information in the **Keywords** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Keywords** property to document the keywords that help locate a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Keywords** property as well as other document properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
Public Sub Keywords_Example()  
 
   Dim vsoDocument As Visio.Document  
 
    Set vsoDocument = Documents.Add("")  
 
    'Set the properties of the document.  
    vsoDocument.Title = "document title "  
    vsoDocument.Creator = "author name "  
    vsoDocument.Description = "document description "  
    vsoDocument.Keywords = "keyword1, keyword2, keyword3 "  
    vsoDocument.Subject = "document subject "  
    vsoDocument.Manager = "manager name "  
    vsoDocument.Category = "document category "  
 
End Sub
```


