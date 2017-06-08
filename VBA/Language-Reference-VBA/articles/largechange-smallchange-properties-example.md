---
title: LargeChange, SmallChange Properties Example
keywords: fm20.chm5225136
f1_keywords:
- fm20.chm5225136
ms.prod: office
ms.assetid: f108dfbc-bf8e-b019-3082-07401e188fbe
ms.date: 06/08/2017
---


# LargeChange, SmallChange Properties Example

The following example demonstrates the  **LargeChange** and **SmallChange** properties when used with a stand-alone **ScrollBar**. The user can set the **LargeChange** and **SmallChange** values to any integer in the range of 0 to 100. This example also uses the **MaxLength** property to restrict the number of characters entered for the **LargeChange** and **SmallChange** values.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Label** named Label1 and a **TextBox** named TextBox1.
    
- A  **Label** named Label2 and a **TextBox** named TextBox2.
    
- A  **ScrollBar** named ScrollBar1.
    
- A  **Label** named Label3.
    




```vb
Dim TempNum As Integer 
 
Private Sub ScrollBar1_Change() 
 Label3.Caption = ScrollBar1.Value 
End Sub 
 
Private Sub TextBox1_Change() 
 If IsNumeric(TextBox1.Text) Then 
 TempNum = CInt(TextBox1.Text) 
 If TempNum >= 0 And TempNum <= 100 Then 
 ScrollBar1.SmallChange = TempNum 
 Else 
 TextBox1.Text = ScrollBar1.SmallChange 
 End If 
 Else 
 TextBox1.Text = ScrollBar1.SmallChange 
 End If 
End Sub 
 
Private Sub TextBox2_Change() 
 If IsNumeric(TextBox2.Text) Then 
 TempNum = CInt(TextBox2.Text) 
 If TempNum >= 0 And TempNum <= 100 Then 
 ScrollBar1.LargeChange = TempNum 
 Else 
 TextBox2.Text = ScrollBar1.LargeChange 
 End If 
 Else 
 TextBox2.Text = ScrollBar1.LargeChange 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 ScrollBar1.Min = -1000 
 ScrollBar1.Max = 1000 
 
 Label1.Caption = "SmallChange 0 to 100" 
 ScrollBar1.SmallChange = 1 
 TextBox1.Text = ScrollBar1.SmallChange 
 TextBox1.MaxLength = 3 
 
 Label2.Caption = "LargeChange 0 to 100" 
 ScrollBar1.LargeChange = 100 
 TextBox2.Text = ScrollBar1.LargeChange 
 TextBox2.MaxLength = 3 
 
 ScrollBar1.Value = 0 
 Label3.Caption = ScrollBar1.Value 
End Sub
```


