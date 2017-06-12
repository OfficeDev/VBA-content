---
title: Document.Subject Property (Visio)
keywords: vis_sdr.chm10514465
f1_keywords:
- vis_sdr.chm10514465
ms.prod: visio
api_name:
- Visio.Document.Subject
ms.assetid: b954ca88-c7f7-0c1f-ed30-8ea3eb3bc0e3
ms.date: 06/08/2017
---


# Document.Subject Property (Visio)

Gets or sets the value of the  **Subject** field in a document's properties. Read/write.


## Syntax

 _expression_ . **Subject**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Subject** property is equivalent to entering information in the **Subject** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Subject** property to document the subject of a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Subject** property as well as other document properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Subject_Example() 
 
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


