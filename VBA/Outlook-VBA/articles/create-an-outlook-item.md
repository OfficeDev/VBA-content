---
title: Create an Outlook Item
keywords: olfm10.chm3077346
f1_keywords:
- olfm10.chm3077346
ms.prod: outlook
ms.assetid: bf84371a-63c2-5b8b-2017-dc01cb79f710
ms.date: 06/08/2017
---


# Create an Outlook Item

This procedure uses the  **Click** event to call **[CreateItem](application-createitem-method-outlook.md)** to create and show an appointment when a user clicks CommandButton1. The example shows how to do this in a custom form page using VBScript.

In design mode:

1. Using the  **Control Toolbox**, place a  **[CommandButton](commandbutton-object-outlook-forms-script.md)** on your form.
    
2. Open the Script Editor. [How](using-the-script-editor.md)?
    
3. Enter the following code, using the appropriate constant value from the  **[OlItemType](olitemtype-enumeration-outlook.md)** enumeration to specify the type of item that you want to create.
    
```vb
  Sub CommandButton1_Click 
 Application.CreateItem(1).Display 
End Sub
```


    
    

