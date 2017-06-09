---
title: "How to: Set the Maximum and Minimum Values for a Scroll Bar"
keywords: olfm10.chm3077214
f1_keywords:
- olfm10.chm3077214
ms.prod: outlook
ms.assetid: 45312fc9-6c40-2dcc-175c-2a64ea635cc8
ms.date: 06/08/2017
---


# How to: Set the Maximum and Minimum Values for a Scroll Bar

The following example demonstrates the  **[Max](scrollbar-max-property-outlook-forms-script.md)** and **[Min](scrollbar-min-property-outlook-forms-script.md)** properties when used with a stand-alone **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. The user can set the  **Max** and **Min** values to any integer in the range of -1000 to 1000. This example also uses the ** [TextBox.MaxLength](textbox-maxlength-property-outlook-forms-script.md)** property to restrict the number of characters entered for the **Max** and **Min** values.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1
    
- A  **[TextBox](textbox-object-outlook-forms-script.md)** named TextBox1 that is bound to the custom number field named ScrollBarMin.
    
- A  **Label** named Label2
    
- A  **TextBox** named TextBox2 that is bound to the custom number field named ScrollBarMax.
    
- A  **ScrollBar** named ScrollBar1 that is bound to the custom number field named ScrollBarValue.
    
- A  **Label** named Label3.
    



```vb
Sub Item_Open() 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set ScrollBar1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ScrollBar1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set Label2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label2") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set Label3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label3") 
 
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
 
Sub Item_CustomPropertyChange(byval pname) 
 Set ScrollBar1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ScrollBar1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set Label3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label3") 
 
 If pname = "ScrollBarMin" Then 
 
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
 ElseIf pname = "ScrollBarMax" Then 
 
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
 ElseIf pname = "ScrollBarValue" Then 
 
 Label3.Caption = ScrollBar1.Value 
 End If 
End Sub
```


