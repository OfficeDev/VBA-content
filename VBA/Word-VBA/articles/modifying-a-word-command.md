---
title: Modifying a Word Command
keywords: vbawd10.chm5212129
f1_keywords:
- vbawd10.chm5212129
ms.prod: word
ms.assetid: bedf22b4-203b-3ecb-1f00-0b88b3bd89e7
ms.date: 06/08/2017
---


# Modifying a Word Command

You can modify most Word commands by turning them into macros. For example, you can modify the  **Open** command on the **File** tab so that instead of displaying a list of Worddocument files (files ending with the .doc file name extension), Word displays every file in the current folder.

To display the list of built-in Word commands in the  **Macro** dialog box (Alt-F8), you select **Word Commands** in the **Macros In** box. Every command available on the ribbon or through shortcut keys is listed. Commands begin with the menu name that was associated with the command before menus were replaced with the ribbon. For example, the **Save** command, which was formerly on the **File** menu, is listed as **FileSave**.

You can replace a Word command with a macro by giving a macro the same name as a Word command. For example, if you create a macro named "FileSave," Word runs the macro when you choose  **Save** from the **File** menu, click the **Save** toolbar button, or press the Ctrl-S shortcut key combination.

This example takes you through the steps needed to modify the FileSave command. 

1. Press Alt-F8.
    
2. In the  **Macros in** box, select **Word commands**.
    
3. In the  **Macro name** box, select "FileSave".
    
4. In the  **Macros in** box, select a template or document location to store the macro. For example, select "Normal.dot (Global Template)" to create a global macro (this modifies the FileSave command for all documents that use the normal tempate).
    
5. Click  **Create**.
    
The FileSave macro appears as shown below.



```vb
Sub FileSave() 
' 
' FileSave Macro 
' Saves the active document or template 
' 
    ActiveDocument.Save 
 
End Sub
```

You can add additional instructions or remove the existing  `ActiveDocument.Save` instruction. Now every time the FileSave command runs, your FileSave macro runs instead of the Word command. To restore the original FileSave functionality, you need to rename or delete your FileSave macro.

## Remarks

You can also replace a Word command by creating a code module named after a Word command (for example, FileSave) with a subroutine named Main.


