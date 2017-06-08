---
title: Document.Title Property (Visio)
keywords: vis_sdr.chm10514540
f1_keywords:
- vis_sdr.chm10514540
ms.prod: visio
api_name:
- Visio.Document.Title
ms.assetid: 9a3b9e5f-2515-dda4-d757-0a0f375dfd00
ms.date: 06/08/2017
---


# Document.Title Property (Visio)

Gets or sets the value of the  **Title** field in a document's properties. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents a **Document** object.


### Return Value

 **String**


## Remarks

Setting the  **Title** property is equivalent to entering information in the **Title** box in the **Properties** dialog box (click the **File** tab, click **Info**, click  **Document Properties**, and then click  **Advanced Properties**).


 **Security Note**  




## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Title** property to set the title of a drawing. It adds a **Document** object to the **Documents** collection and sets the **Document** object's **Title** property as well as other document properties.

Before running this macro, substitute your own values for the items in italic in the following code. To verify that these properties have been set, open the  **Properties** dialog box.




```vb
 
Public Sub Title_Example() 
  
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


