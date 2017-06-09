---
title: Use the PropertyChange Event to Test Procedures
keywords: olfm10.chm3077359
f1_keywords:
- olfm10.chm3077359
ms.prod: outlook
ms.assetid: 9e0beb04-dc64-ad5d-ae77-8c11c11349b0
ms.date: 06/08/2017
---


# Use the PropertyChange Event to Test Procedures

This topic shows how to test procedures for a custom form page that uses VBScript and the Script Editor.

Perform the following steps to test simple procedures. Replace the code below with the code that you want to test. Each time a user changes the value of the Importance field, or any other default field, the code runs.

1. Open the Script Editor. [How](using-the-script-editor.md)?
    
2. On the  **Script** menu, click **Event Handler**.
    
3. In the  **Events** box, double-click **PropertyChange**.
    
4. Add the following code in the event:
    
```vb
  MsgBox "This is my test procedure"
```


    
    
5. On the  **Form** menu, click **Run This Form**.
    
6. Click the  **!** icon on the toolbar. The message box will appear.
    
7. Click  **OK** to close the message box.
    

