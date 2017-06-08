---
title: Max, Min, MaxLength Properties Example
keywords: fm20.chm5225135
f1_keywords:
- fm20.chm5225135
ms.prod: office
ms.assetid: 17886973-605e-3fc6-5df4-677355932c14
ms.date: 06/08/2017
---


# Max, Min, MaxLength Properties Example

The following example demonstrates the  **Max** and **Min** properties when used with a stand-alone **ScrollBar**. The user can set the **Max** and **Min** values to any integer in the range of -1000 to 1000. This example also uses the **MaxLength** property to restrict the number of characters entered for the **Max** and **Min** values.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Label** named Label1 and a **TextBox** named TextBox1.
    
- A  **Label** named Label2 and a **TextBox** named TextBox2.
    
- A  **ScrollBar** named ScrollBar1.
    
- A  **Label** named Label3.
    




```vb
Dim TempNum As Integer 
 
Private Sub UserForm_Initialize() 
 Label1.Caption = "Min -1000 to 1000" 
 ScrollBar1.Min = -1000 
 TextBox1.Text = ScrollBar1.Min 
 TextBox1.MaxLength = 5 
 
 Label2.Caption = "Max -1000 to 1000" 
 ScrollBar1.Max = 1000 
 TextBox2.Text = ScrollBar1.Max 
 TextBox2.MaxLength = 5 
 
 ScrollBar1.SmallChange = 1 
 ScrollBar1.LargeChange = 100 
 ScrollBar1.Value = 0 
 Label3.Caption = ScrollBar1.Value 
End Sub 
 
Private Sub TextBox1_Change() 
 If IsNumeric(TextBox1.Text) Then 
 TempNum = CInt(TextBox1.Text) 
 If TempNum >= -1000 And TempNum <= 1000 Then 
 ScrollBar1.Min = TempNum 
 Else 
 TextBox1.Text = ScrollBar1.Min 
 End If 
 Else 
 TextBox1.Text = ScrollBar1.Min 
 End If 
End Sub 
 
Private Sub TextBox2_Change() 
 If IsNumeric(TextBox2.Text) Then 
 TempNum = CInt(TextBox2.Text) 
 If TempNum >= -1000 And TempNum <= 1000 Then 
 ScrollBar1.Max = TempNum 
 Else 
 TextBox2.Text = ScrollBar1.Max 
 End If 
 Else 
 TextBox2.Text = ScrollBar1.Max 
 End If 
End Sub 
 
Private Sub ScrollBar1_Change() 
Label3.Caption = ScrollBar1.Value 
End Sub
```


