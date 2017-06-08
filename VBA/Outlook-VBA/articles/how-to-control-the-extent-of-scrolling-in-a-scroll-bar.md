---
title: "How to: Control the Extent of Scrolling in a Scroll Bar"
keywords: olfm10.chm3077200
f1_keywords:
- olfm10.chm3077200
ms.prod: outlook
ms.assetid: 60a8eea3-9277-4db0-ffa8-5ad2a8adb0b4
ms.date: 06/08/2017
---


# How to: Control the Extent of Scrolling in a Scroll Bar

The following example demonstrates the  **[LargeChange](scrollbar-largechange-property-outlook-forms-script.md)** and **[SmallChange](scrollbar-smallchange-property-outlook-forms-script.md)** properties when used with a stand-alone **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. The user can set the  **LargeChange** and **SmallChange** values to any integer in the range of 0 to 100. This example also uses the ** [TextBox.MaxLength](textbox-maxlength-property-outlook-forms-script.md)** property to restrict the number of characters entered in the **[TextBox](textbox-object-outlook-forms-script.md)** controls for the **LargeChange** and **SmallChange** values.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1
    
- A  **TextBox** named TextBox1 that is bound to the custom number field named ScrollBarSmallChange
    
- A  **Label** named Label2
    
- A  **TextBox** named TextBox2 that is bound to the custom number field named ScrollBarLargeChange.
    
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
 
Sub Item_CustomPropertyChange(byval pname) 
 Set ScrollBar1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ScrollBar1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set Label3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label3") 
 
 If pname = "ScrollBarMin" Then 
 
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
 
 ElseIf pname = "ScrollBarMax" Then 
 
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
 
 ElseIf pname = "ScrollBarValue" Then 
 
 Label3.Caption = ScrollBar1.Value 
 End If 
End Sub
```


