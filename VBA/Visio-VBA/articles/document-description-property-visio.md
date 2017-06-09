---
title: Document.Description Property (Visio)
keywords: vis_sdr.chm10513405
f1_keywords:
- vis_sdr.chm10513405
ms.prod: visio
api_name:
- Visio.Document.Description
ms.assetid: 530adbe3-5285-6aa5-32e6-88d2bc1d8ebf
ms.date: 06/08/2017
---


# Document.Description Property (Visio)

Gets or sets the description of a document, one of a document's properties. Read/write.


## Syntax

 _expression_ . **Description**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting a document's  **Description** property is equivalent to entering information in the **Description** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Description** property to document the description of a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Description** property as well as other document properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Description_Example()  
 
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


