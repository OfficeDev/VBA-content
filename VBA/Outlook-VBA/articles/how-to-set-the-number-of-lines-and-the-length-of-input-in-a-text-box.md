---
title: "How to: Set the Number of Lines and the Length of Input in a Text Box"
keywords: olfm10.chm3077204
f1_keywords:
- olfm10.chm3077204
ms.prod: outlook
ms.assetid: 1b56aff7-ab6f-b595-781d-a60d0dffe7a9
ms.date: 06/08/2017
---


# How to: Set the Number of Lines and the Length of Input in a Text Box

The following example counts the characters and the number of lines of text in a  **[TextBox](textbox-object-outlook-forms-script.md)** by using the **[LineCount](textbox-linecount-property-outlook-forms-script.md)** and **[TextLength](textbox-textlength-property-outlook-forms-script.md)** properties, and the **SetFocus** method. In this example, the user can type into a **TextBox**, and can retrieve current values of the  **LineCount** and **TextLength** properties.


 **Note**  The  **SetFocus** method is inherited from the Microsoft Forms 2.0 **TextBox** control.


To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains the following controls:


- A  **TextBox** named TextBox1.
    
- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    
- Two  **[Label](label-object-outlook-forms-script.md)** controls named Label1 and Label2.
    



```vb
'Type SHIFT+ENTER to start a new line in the text box. 
 
Dim CommandButton1 
Dim TextBox1 
Dim Label1 
Dim Label2 
 
Sub CommandButton1_Click() 
 'Must first give TextBox1 the focus to get line count 
 TextBox1.SetFocus 
 Label1.Caption = "LineCount = " &; TextBox1.LineCount 
 Label2.Caption = "TextLength = " &; TextBox1.TextLength 
End Sub 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("Label1") 
 Set Label2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("Label2") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton1") 
 
 CommandButton1.WordWrap = True 
 CommandButton1.AutoSize = True 
 CommandButton1.Caption = "Get Counts" 
 
 Label1.Caption = "LineCount = " 
 Label2.Caption = "TextLength = " 
 
 TextBox1.MultiLine = True 
 TextBox1.WordWrap = True 
 TextBox1.Text = "Enter your text here." 
End Sub
```


