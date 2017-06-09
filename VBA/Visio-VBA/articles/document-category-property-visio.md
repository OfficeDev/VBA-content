---
title: Document.Category Property (Visio)
keywords: vis_sdr.chm10513175
f1_keywords:
- vis_sdr.chm10513175
ms.prod: visio
api_name:
- Visio.Document.Category
ms.assetid: da312b56-6232-9077-e47b-47144aa603c5
ms.date: 06/08/2017
---


# Document.Category Property (Visio)

Gets or sets the value of a document's category, one of the document properties. Read/write.


## Syntax

 _expression_ . **Category**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Category** property is equivalent to entering information in the **Categories** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  ** Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic macro shows how to use the  **Category** property to categorize a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Category** property as well as other properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Category_Example() 
  
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


