---
title: Document.Creator Property (Visio)
keywords: vis_sdr.chm10513335
f1_keywords:
- vis_sdr.chm10513335
ms.prod: visio
api_name:
- Visio.Document.Creator
ms.assetid: c1dea222-796c-1231-bb9b-9d258450b142
ms.date: 06/08/2017
---


# Document.Creator Property (Visio)

Gets or sets the value of a document's author—one of the document's properties. Read/write.


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Creator** property is equivalent to entering information in the **Author** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Creator** property to document the author of a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Creator** property as well as other document properties.

 Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the **Properties** dialog box.




```vb
 
Public Sub Creator_Example() 
    
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


