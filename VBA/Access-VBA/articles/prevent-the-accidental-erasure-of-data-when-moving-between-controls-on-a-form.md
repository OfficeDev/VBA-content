---
title: Prevent the Accidental Erasure of Data When Moving Between Controls on a Form
ms.prod: access
ms.assetid: 1733caa5-5067-e6d9-b614-51053180f22e
ms.date: 06/08/2017
---


# Prevent the Accidental Erasure of Data When Moving Between Controls on a Form

When you tab from one text box or memo field to another in a form, the text in the control is highlighted. This makes it easy for users to accidentally delete the text by pressing a key. By using a few lines of code, you can move the insertion point to the first position in the text box, minimizing the risk of accidentally deleting the text. 

To do this, create a procedure for the text box's  **[GotFocus](textbox-gotfocus-event-access.md)** event. In the **GotFocus** event procedure, set the **[SelLength](textbox-sellength-property-access.md)** property of the text box to its **[SelStart](combobox-selstart-property-access.md)** property. The following example illustrates how to do this for a text box named **txtFirstName**.



```vb
Private Sub txtFirstName_GotFocus() 
 
    Me.txtFirstName.SelLength = Me.txtFirstName.SelStart 
 
End Sub
```


