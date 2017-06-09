---
title: "How to: Undo and Redo an Action in a Text Box or a Combo Box"
keywords: olfm10.chm3077161
f1_keywords:
- olfm10.chm3077161
ms.prod: outlook
ms.assetid: 5edb515c-d035-e8a3-8c0b-f2ddc74378fd
ms.date: 06/08/2017
---


# How to: Undo and Redo an Action in a Text Box or a Combo Box

The following example demonstrates how to undo or redo text editing within a  **[TextBox](textbox-object-outlook-forms-script.md)** or within the text area of a **[ComboBox](combobox-object-outlook-forms-script.md)**. This sample checks whether an undo or redo operation can occur and then performs the appropriate action. The sample uses the  **CanUndo** and **CanRedo** properties, and the **UndoAction** and **RedoAction** methods.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **TextBox** named TextBox1.
    
- A  **ComboBox** named ComboBox1.
    
- Two  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls named CommandButton1 and CommandButton2.
    



```vb
Dim UserForm1 
 
Sub CommandButton1_Click() 
 If UserForm1.CanUndo = True Then 
 UserForm1.UndoAction 
 MsgBox "Undid IT" 
 Else 
 MsgBox "No undo performed." 
 End If 
End Sub 
 
Sub CommandButton2_Click() 
 If UserForm1.CanRedo = True Then 
 UserForm1.RedoAction 
 MsgBox "Redid IT" 
 Else 
 MsgBox "No redo performed." 
 End If 
End Sub 
 
Sub Item_Open() 
 Set UserForm1 = Item.GetInspector.ModifiedFormPages("P.2") 
 Set TextBox1 = UserForm1.Controls("TextBox1") 
 Set ComboBox1 = UserForm1.Controls("ComboBox1") 
 Set CommandButton1 = UserForm1.Controls("CommandButton1") 
 Set CommandButton2 = UserForm1.Controls("CommandButton2") 
 
 TextBox1.Text = "Type your text here." 
 
 ComboBox1.ColumnCount = 3 
 ComboBox1.AddItem "Choice 1, column 1" 
 ComboBox1.List(0, 1) = "Choice 1, column 2" 
 ComboBox1.List(0, 2) = "Choice 1, column 3" 
 
 CommandButton1.Caption = "Undo" 
 CommandButton2.Caption = "Redo" 
End Sub
```


