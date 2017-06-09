---
title: Hyperlink.SubAddress Property (Access)
keywords: vbaac10.chm10114
f1_keywords:
- vbaac10.chm10114
ms.prod: access
api_name:
- Access.Hyperlink.SubAddress
ms.assetid: b281fa9e-502b-59b4-749e-3c96913e4d14
ms.date: 06/08/2017
---


# Hyperlink.SubAddress Property (Access)

You can use the  **SubAddress** property to specify or determine a location within the target document specified by the **[Address](hyperlink-address-property-access.md)** property. Read/write **String**. .


## Syntax

 _expression_. **SubAddress**

 _expression_ A variable that represents a **Hyperlink** object.


## Remarks

The  **SubAddress** property can be an object within a Microsoft Access database, a bookmark within a Microsoft Word document, a named range within a Microsoft Excel spreadsheet, a slide within a Microsoft PowerPoint presentation, or a location within an HTML document.

The  **SubAddress** property represents the **HyperlinkSubAddress** property of a named location within the target document specified by the **HyperlinkAddress** property

When you move the cursor over a command button, image control, or label control whose  **HyperlinkSubAddress** property is set, the cursor changes to an upward-pointing hand. Clicking the control displays the object or Web page specified by the link.

For more information about hyperlink addresses and their format, see the  **HyperlinkAddress** and **HyperlinkSubAddress** property topics.


## Example

The following example turns a label named "Label20" on the "Suppliers" form into an active hyperlink. When the user click the hyperlink, Access opens the "Mailing List" form in the "Postal Operations" database.


```vb
With Forms("Suppliers").Controls("Label20").Hyperlink 
 .Address = "PostalOperations.mdb" 
 .SubAddress = "Form Mailing List" 
End With
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

