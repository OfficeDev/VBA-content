---
title: Recording a Macro to Generate Code
keywords: vbawd10.chm5212835
f1_keywords:
- vbawd10.chm5212835
ms.prod: word
ms.assetid: cd41b71f-567b-d156-1d5e-973e276f27fb
ms.date: 06/08/2017
---


# Recording a Macro to Generate Code

If you are unsure of which Visual Basic method or property to use, you can turn on the macro recorder and manually perform the action. The macro recorder translates your actions into Visual Basic code. After you record your actions, you can modify the code to do exactly what you want. For example, if you do not know what property or method to use to indent a paragraph, do the following:


1. On the  **Developer** ribbon, click **Record Macro**.
    
2. Change the default macro name to a name of your choice and click  **OK** to start the recorder.
    
3. On the  **Home** menu, click the **Increase Indent** button.
    
4. On the  **Developer** ribbon, click **Stop Recording**.
    
5. On the  **Developer** ribbon, click **Macros**.
    
6. Select the macro name that you assigned in Step 2 and click  **Edit**.
    

View the Visual Basic code to determine the property that corresponds to the left paragraph indent (the  **[LeftIndent](paragraph-leftindent-property-word.md)** property). Position the cursor within `.LeftIndent` and press **F1** or click the **Help** button.


## Remarks

Recorded macros use the  **[Selection](selection-object-word.md)** object. The following code example indents the selected paragraphs by one-half inch.


```vb
Sub IndentParagraph() 
    Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5) 
End Sub
```

You can, however, modify the recorded macro to work with  **[Range](range-object-word.md)** objects. For more information, see [Revising Recorded Visual Basic Macros](revising-recorded-visual-basic-macros.md).


