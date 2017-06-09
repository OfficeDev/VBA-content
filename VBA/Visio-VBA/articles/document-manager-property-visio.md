---
title: Document.Manager Property (Visio)
keywords: vis_sdr.chm10513865
f1_keywords:
- vis_sdr.chm10513865
ms.prod: visio
api_name:
- Visio.Document.Manager
ms.assetid: 6aa5bcfc-55b5-88ce-a9a8-d0f6a73ee69f
ms.date: 06/08/2017
---


# Document.Manager Property (Visio)

Gets or sets the value of the  **Manager** box in a document's **Properties** dialog box. Read/write.


## Syntax

 _expression_ . **Manager**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Manager** property is equivalent to entering information in the **Manager** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Manager** property to document the name of the manager of the author of a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Manager** property as well as other document properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Manager_Example() 
  
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


