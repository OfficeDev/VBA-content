---
title: "How to: Display the Name of Each Control on a Form or a Page of a MultiPage Control"
keywords: olfm10.chm3077199
f1_keywords:
- olfm10.chm3077199
ms.prod: outlook
ms.assetid: 503b16dd-51d8-450b-fa1f-0e114a3b9b04
ms.date: 06/08/2017
---


# How to: Display the Name of Each Control on a Form or a Page of a MultiPage Control

The following example uses the  **Item** method to access individual members of the Microsoft Forms 2.0 **Controls** collection and **[Pages](pages-object-outlook-forms-script.md)** collection. The user chooses an option button for either the **Controls** collection or the **[MultiPage](multipage-object-outlook-forms-script.md)**, and then clicks the  **[CommandButton](commandbutton-object-outlook-forms-script.md)**. The name of the appropriate control is returned in the  **[Label](label-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **CommandButton** named CommandButton1.
    
- A  **Label** named Label1.
    
- Two  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls named OptionButton1 and OptionButton2.
    
- A  **MultiPage** named MultiPage1.
    



```vb
Dim ControlsIndex 
 
Sub CommandButton1_Click() 
 Set Controls = Item.GetInspector.ModifiedFormPages("P.2").Controls 
 Set OptionButton1 = Controls("OptionButton1") 
 Set OptionButton2 = Controls("OptionButton2") 
 Set Label1 = Controls("Label1") 
 Set MultiPage1 = Controls("MultiPage1") 
 
 If OptionButton1.Value = True Then 
 'Process Controls collection for UserForm 
 Set MyControl = Controls.Item(ControlsIndex) 
 Label1.Caption = MyControl.Name 
 
 'Prepare index for next control on Userform 
 ControlsIndex = ControlsIndex + 1 
 If ControlsIndex >= Controls.Count Then 
 ControlsIndex = 0 
 End If 
 
 ElseIf OptionButton2.Value = True Then 
 'Process Current Page of Pages collection 
 Set MyControl = MultiPage1.Pages.Item(MultiPage1.Value) 
 Label1.Caption = MyControl.Name 
 End If 
End Sub 
 
Sub Item_Open() 
 ControlsIndex = 0 
 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 OptionButton1.Caption = "Controls Collection" 
 OptionButton2.Caption = "Pages Collection" 
 OptionButton1.Value = True 
 
 CommandButton1.Caption = "Get Member Name" 
End Sub
```


