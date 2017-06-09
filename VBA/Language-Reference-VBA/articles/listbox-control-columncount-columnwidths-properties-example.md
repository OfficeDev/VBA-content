---
title: ListBox Control, ColumnCount, ColumnWidths Properties Example
keywords: fm20.chm5225185
f1_keywords:
- fm20.chm5225185
ms.prod: office
ms.assetid: f2f6e0f7-504d-1565-4dcb-d8bd2ff129c7
ms.date: 06/08/2017
---


# ListBox Control, ColumnCount, ColumnWidths Properties Example

The following example uses the  **ColumnWidths** property to change the column widths of a multicolumn **ListBox**. The example uses three **TextBox** controls to specify the individual column widths and uses the **Exit** event to specify the units of measure of each **TextBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **ListBox** named ListBox1.
    
- Three  **TextBox** controls named TextBox1 through TextBox3.
    
- A  **CommandButton** named CommandButton1.
    

Try entering the value 0 to hide a column.



```vb
Dim MyArray(2, 3) As String 
 
Private Sub CommandButton1_Click() 
 'ColumnWidths requires a value for each column 
 'separated by semicolons 
 ListBox1.ColumnWidths = TextBox1.Text &; ";" _ 
 &; TextBox2.Text &; ";" &; TextBox3.Text 
End Sub
```




```vb
Private Sub TextBox1_Exit(ByVal Cancel As _ 
 MSForms.ReturnBoolean) 
 'ColumnWidths accepts points (no units), inches 
 'or centimeters; make inches the default 
 If Not (InStr(TextBox1.Text, "in") > 0 Or _ 
 InStr(TextBox1.Text, "cm") > 0) Then 
 TextBox1.Text = TextBox1.Text &; " in" 
 End If 
End Sub
```




```vb
Private Sub TextBox2_Exit(ByVal Cancel As _ 
 MSForms.ReturnBoolean) 
 'ColumnWidths accepts points (no units), inches 
 'or centimeters; make inches the default 
 If Not (InStr(TextBox2.Text, "in") > 0 Or _ 
 InStr(TextBox2.Text, "cm") > 0) Then 
 TextBox2.Text = TextBox2.Text &; " in" 
 End If 
End Sub
```




```vb
Private Sub TextBox3_Exit(ByVal Cancel as MSForms.ReturnBoolean) 
 'ColumnWidths accepts points (no units), inches or 
 'centimeters; make inches the default 
 If Not (InStr(TextBox3.Text, "in") > 0 Or _ 
 InStr(TextBox3.Text, "cm") > 0) Then 
 TextBox3.Text = TextBox3.Text &; " in" 
 End If 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
Dim i, j, Rows As Single 
 
ListBox1.ColumnCount = 3 
Rows = 2 
 
For j = 0 To ListBox1.ColumnCount - 1 
 For i = 0 To Rows - 1 
 MyArray(i, j) = "Row " &; i &; ", Column " &; j 
 Next i 
Next j 
'Load MyArray into ListBox1 
ListBox1.List() = MyArray 
'1-inch columns initially 
TextBox1.Text = "1 in" 
TextBox2.Text = "1 in" 
TextBox3.Text = "1 in" 
End Sub
```


