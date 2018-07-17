---
title: "How to: Change the Column Widths of a Multi-Column List Box"
keywords: olfm10.chm3077166
f1_keywords:
- olfm10.chm3077166
ms.prod: outlook
ms.assetid: ad79a33e-ec14-0f37-468b-de1b833f1f84
ms.date: 06/08/2017
---


# How to: Change the Column Widths of a Multi-Column List Box

The following example uses the  **[ColumnWidths](listbox-columnwidths-property-outlook-forms-script.md)** property to change the column widths of a multicolumn **[ListBox](listbox-object-outlook-forms-script.md)**. The example uses three  **[TextBox](textbox-object-outlook-forms-script.md)** controls to specify the individual column widths and uses the **[Click](commandbutton-click-event-outlook-forms-script.md)** event to specify the units of measure of each **TextBox**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ListBox** named ListBox1.
    
- Three custom text fields named Text1, Text2, and Text3.
    
- Three  **TextBox** controls named TextBox1, TextBox2, and TextBox3 that are bound to the custom text fields above.
    
- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    
Try entering the value 0 to hide a column.



```vb
Dim MyArray(2, 3) 
Dim ListBox1 
Dim TextBox1 
Dim TextBox2 
Dim TextBox3 
Dim CommandButton1 
 
Sub Item_Open() 
Dim i, j, Rows 
 
Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").ListBox1 
Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").TextBox1 
Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").TextBox2 
Set TextBox3 = Item.GetInspector.ModifiedFormPages("P.2").TextBox3 
Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 
ListBox1.ColumnCount = 3 
Rows = 2 
 
For j = 0 To ListBox1.ColumnCount - 1 
 For i = 0 To Rows - 1 
 MyArray(i, j) = "Row " &; i &; ", Column " &; j 
 Next 
Next 
 
ListBox1.List() = MyArray 'Load MyArray into ListBox1 
 
TextBox1.Text = "1 in" '1-inch columns initially 
TextBox2.Text = "1 in" 
TextBox3.Text = "1 in" 
 
End Sub 
 
Sub CommandButton1_Click() 
 'ColumnWidths requires a value for each column separated by semicolons 
 ListBox1.ColumnWidths = TextBox1.Text &; ";" &; TextBox2.Text &; ";" &; TextBox3.Text 
End Sub 
 
Sub Item_CustomPropertyChange(ByVal Name) 
msgbox Name 
Select Case Name 
Case "Text1" 
 'ColumnWidths accepts points (no units), inches or centimeters; make inches the default 
 If Not (InStr(TextBox1.Text, "in") > 0 Or InStr(TextBox1.Text, "cm") > 0) Then 
 TextBox1.Text = TextBox1.Text &; " in" 
 End If 
Case "Text2" 
 'ColumnWidths accepts points (no units), inches or centimeters; make inches the default 
 If Not (InStr(TextBox2.Text, "in") > 0 Or InStr(TextBox2.Text, "cm") > 0) Then 
 TextBox2.Text = TextBox2.Text &; " in" 
 End If 
Case "Text3" 
 'ColumnWidths accepts points (no units), inches or centimeters; make inches the default 
 If Not (InStr(TextBox3.Text, "in") > 0 Or InStr(TextBox3.Text, "cm") > 0) Then 
 TextBox3.Text = TextBox3.Text &; " in" 
 End If 
End Select 
End Sub
```


