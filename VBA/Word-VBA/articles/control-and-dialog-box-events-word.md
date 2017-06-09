---
title: Control and Dialog Box Events (Word)
keywords: vbawd10.chm5210327
f1_keywords:
- vbawd10.chm5210327
ms.prod: word
ms.assetid: 7884bae3-caa5-79a9-a4a2-c58a6ccb42d2
ms.date: 06/08/2017
---


# Control and Dialog Box Events (Word)

After you have added  [ActiveX controls](http://msdn.microsoft.com/library/befa20c2-c4e7-1a53-7740-248885691710%28Office.15%29.aspx)to your dialog box or document, you add event procedures to determine how the controls respond to user actions.

UserForms and controls have a predefined set of events. For example, a command button has a  **Click** event that occurs when the user clicks the command button, and UserForms have an **Initialize** event that runs when the form is loaded.

To write a control or form event procedure, open a module by double-clicking the form or control, and then select the event from the  **Procedure** drop-down list box.

Event procedures include the name of the control. For example, the name of the  **Click** event procedure for a command button named Command1 is Command1_Click.
If you add code to an event procedure and then change the name of the control, your code remains in procedures with the previous name.
For example, assume you add code to the  **Click** event for Commmand1 and then rename the control to Command2. 

When you double-click Command2, you will not see any code in the **Click** event procedure. You will need to move code from Command1_Click to Command2_Click.
To simplify development, it is a good practice to name your controls before writing code.

