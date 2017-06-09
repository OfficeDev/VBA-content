---
title: Application.Echo Method (Access)
keywords: vbaac10.chm12505
f1_keywords:
- vbaac10.chm12505
ms.prod: access
api_name:
- Access.Application.Echo
ms.assetid: ce94d774-ef06-7cf4-0e91-b5affa41a437
ms.date: 06/08/2017
---


# Application.Echo Method (Access)

The  **Echo** method specifies whether Microsoft Access repaints the display screen.


## Syntax

 _expression_. **Echo**( ** _EchoOn_**, ** _bstrStatusBarText_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EchoOn_|Required|**Integer**|**True** (default) indicates that the screen is repainted.|
| _bstrStatusBarText_|Optional|**String**|A string expression that specifies the text to display in the status bar when repainting is turned on or off.|

## Remarks

If you are running Visual Basic code that makes a number of changes to objects displayed on the screen, your code may work faster if you turn off screen repainting until the procedure has finished running. You may also want to turn repainting off if your code makes changes that the user should not or does not need to see.

The  **Echo** method does not suppress the display of modal dialog boxes, such as error messages, or pop-up forms, such as property sheets.


 **Note**  The  **Echo** method doesn't affect the visibility of the ribbon or the availability of ribbon commands.

If you turn screen repainting off, the screen won't show any changes, even if the user presses CTRL+BREAK or Visual Basic encounters a breakpoint. You may want to create a macro that turns repainting on and then assign the macro to a key or custom menu command. You can then use the key combination or menu command to turn repainting on if it has been turned off in Visual Basic.

If you turn screen repainting off and then try to step through the code, you won't be able to see progress through the code or any other visual cues until repainting is turned back on. However, your code will continue to execute.


 **Note**  Do not confuse the  **Echo** method with the **Repaint** method. The **Echo** method turns screen repainting on or off. The **Repaint** method forces an immediate screen repainting.


## Example

The following code example uses the  **Echo** method to prevent the screen from being repainted while certain operations are underway. While the procedure opens a form and minimizes it, the user only sees an hourglass icon indicating that processing is taking place, and the screen isn't repainted. When this task is completed, the hourglass changes back to a pointer and screen repainting is turned back on.


```vb
Public Sub EchoOff() 
 
 ' Open the Employees form minimized. 
 Application.Echo False 
 DoCmd.Hourglass True 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.Minimize 
 Application.Echo True 
 DoCmd.Hourglass False 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

