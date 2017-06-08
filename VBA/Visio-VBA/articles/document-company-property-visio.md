---
title: Document.Company Property (Visio)
keywords: vis_sdr.chm10513285
f1_keywords:
- vis_sdr.chm10513285
ms.prod: visio
api_name:
- Visio.Document.Company
ms.assetid: b55e23dc-3b58-c062-1738-74d2f50fa39d
ms.date: 06/08/2017
---


# Document.Company Property (Visio)

Gets or sets the name of the company the document belongs to, one of the document's properties. Read/write.


## Syntax

 _expression_ . **Company**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Setting the  **Company** property is equivalent to entering information in the **Company** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic macro shows how to use the  **Company** property to document the company for which the drawing is made. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Company** property as well as other properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Company_Example() 
  
    Dim vsoDocument As Visio.Document  
 
    Set vsoDocument = Documents.Add("")  
 
    'Set the properties of the document.  
    vsoDocument.Title = "document title "  
    vsoDocument.Creator = "author name "  
    vsoDocument.Description = "document description "  
    vsoDocument.Keywords = "keyword1, keyword2, keyword3 "  
    vsoDocument.Subject = "document subject "  
    vsoDocument.Manager = "manager name "  
    vsoDocument.Company = "company name "  
 
End Sub
```


