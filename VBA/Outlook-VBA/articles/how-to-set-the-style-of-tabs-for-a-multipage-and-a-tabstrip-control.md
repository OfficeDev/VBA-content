---
title: "How to: Set the Style of Tabs for a MultiPage and a TabStrip Control"
keywords: olfm10.chm3077246
f1_keywords:
- olfm10.chm3077246
ms.prod: outlook
ms.assetid: abaeceb2-d8fd-92a5-8d54-eab94d9de394
ms.date: 06/08/2017
---


# How to: Set the Style of Tabs for a MultiPage and a TabStrip Control

The following example uses the  **Style** property to specify the appearance of the tabs in **[MultiPage](multipage-object-outlook-forms-script.md)** and **[TabStrip](tabstrip-object-outlook-forms-script.md)**. This example also demonstrates using a  **[Label](label-object-outlook-forms-script.md)**. The user chooses a style by selecting an  **[OptionButton](optionbutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **Label** named Label1.
    
- Three  **OptionButton** controls named OptionButton1 through OptionButton3.
    
- A  **MultiPage** named MultiPage1.
    
- A  **TabStrip** named TabStrip1.
    
- Any control inside the  **TabStrip**.
    
- Any control in each page of the  **MultiPage**.
    



```vb
Sub OptionButton1_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 0 '0=fmTabStyleTabs 
 TabStrip1.Style = 0 '0=fmTabStyleTabs 
End Sub 
 
Sub OptionButton2_Click() 
 'Note that the page borders are invisible 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 1 '1=fmTabStyleButtons 
 TabStrip1.Style = 1 '1=fmTabStyleButtons 
End Sub 
 
Sub OptionButton3_Click() 
 'Note that the page borders are invisible and 
 'the page body begins where the tabs normally appear. 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 2 '2=fmTabStyleNone 
 TabStrip1.Style = 2 '2=fmTabStyleNone 
End Sub 
 
Sub Item_Open() 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 
 Label1.Caption = "Page/Tab Style" 
 OptionButton1.Caption = "Tabs" 
 OptionButton1.Value = True 
 MultiPage1.Style = 0 '0=fmTabStyleTabs 
 TabStrip1.Style = 0 '0=fmTabStyleTabs 
 
 OptionButton2.Caption = "Buttons" 
 OptionButton3.Caption = "No Tabs or Buttons" 
End Sub
```


