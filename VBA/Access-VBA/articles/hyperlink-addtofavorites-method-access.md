---
title: Hyperlink.AddToFavorites Method (Access)
keywords: vbaac10.chm10116
f1_keywords:
- vbaac10.chm10116
ms.prod: access
api_name:
- Access.Hyperlink.AddToFavorites
ms.assetid: 42c92d93-68a1-a5b9-8517-e9257100434e
ms.date: 06/08/2017
---


# Hyperlink.AddToFavorites Method (Access)

The  **AddToFavorites** method adds a hyperlink address to the Favorites folder.


## Syntax

 _expression_. **AddToFavorites**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

Nothing


## Remarks

When applied to a  **Control** object, the **AddToFavorites** method adds the hyperlink address contained in a control to the Favorites folder.


## Example

The following example sets the  **HyperlinkAddress** property of a command button. When the user clicks the command button, the address is added to the Favorites folder by using the **AddToFavorites** method.

To try this example, create a new form and add a command button named Command0. Paste the following code into the form's module. Switch to Form view and click the command button.




```vb
Private Sub Form_Load() 
 Me!Command0.HyperlinkAddress = "http://www.microsoft.com/" 
End SubPrivate Sub Command0_Click() 
 Me!Command0.Hyperlink.AddToFavorites 
End Sub
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

