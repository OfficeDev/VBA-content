---
title: Using ActiveX Controls on Word Documents
ms.prod: word
ms.assetid: 529119ff-9108-70cf-d692-ec1fbb37e157
ms.date: 06/08/2017
---


# Using ActiveX Controls on Word Documents

You can add controls to your documents to create interactive documents, such as online forms. For general information about adding and working with controls, see  [Using ActiveX controls on a document](using-activex-controls-on-a-document-word.md) and [Creating a custom dialog box](creating-a-custom-dialog-box.md).

Remember the following points when you work with controls on documents:

- You can add ActiveX controls to the text layer or drawing layer of the document. To add a control to the drawing layer, click the control on the  **Control Toolbox**. To add a control to the text layer, hold down the SHIFT key while you click a control on the  **Control Toolbox**.
    
- A control that you add to the text layer is an  **[InlineShape](inlineshape-object-word.md)** object, to which you gain access programmatically through the **[InlineShapes](inlineshapes-object-word.md)** collection. A control that you add to the drawing layer is a **[Shape](shape-object-word.md)** object, to which you gain access programmatically through the **[Shapes](shapes-object-word.md)** collection.
    
- Controls in the text layer are treated like characters and are positioned as characters within a line of text.
    
- In design mode, ActiveX controls in the drawing layer are visible only in print layout view or Web layout view.
    
- If you want the user to use the ActiveX controls but not change the layout of the document, protect the document by clicking the  **Protect Form** button on the **Forms** toolbar.
    
- Microsoft Word implements the  **LostFocus** and **GotFocus** events for ActiveX controls on a document. The other events listed in the **Procedure** drop-down list box are documented in Microsoft Forms Help. For more information about using events with ActiveX controls, see the [Control and dialog box events](control-and-dialog-box-events-word.md) and the [Using events with ActiveX controls](using-events-with-activex-controls.md) topics.
    
- If you want to add form fields instead of ActiveX controls to your document to create an online form, use the  **Forms** toolbar.
    
- The  **Me** keyword in an event procedure for an ActiveX control on a document refers to the document, not to the control.
    
Writing event code for controls on documents is very similar to writing event code for controls on forms. The following  **SpinUp** and **SpinDown** event procedures change the value of the **TextBox** control named "TextBox1" on the document where the **SpinButton** control named "SpinButton1" resides. The text box value is decreased by one when the user clicks the lower spin-button arrow or the left spin-button arrow and is incremented by one when the user clicks the upper spin-button arrow or the right spin-button arrow.



```vb
Private Sub SpinButton1_SpinDown() 
 Me.TextBox1.Value = Me.TextBox1.Value - 1 
End Sub
```




```vb
Private Sub SpinButton1_SpinUp() 
 Me.TextBox1.Value = Me.TextBox1.Value + 1 
End Sub
```

The following  **Click** event procedure switches to print view and sets the magnification to 100 percent for the document where the command button named "cmdChangeView" resides.



```vb
Private Sub cmdChangeView_Click() 
 With Me.ActiveWindow.View 
 .Type = wdPrintView 
 .Zoom.Percentage = 100 
 End With 
End Sub
```


