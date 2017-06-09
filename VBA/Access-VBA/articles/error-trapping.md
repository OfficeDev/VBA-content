---
title: Error Trapping
keywords: vbaac10.chm5186626
f1_keywords:
- vbaac10.chm5186626
ms.prod: access
ms.assetid: 41d8de92-55ed-8537-eb31-6d72ba69c165
ms.date: 06/08/2017
---


# Error Trapping

You can use the  **On Error GoTo** statement to trap errors and direct procedure flow to the location of error-handling statements within a procedure. For example, the following statement directs the flow to the label line:


```vb
On Error GoTo ErrorHandler
```


Be sure to give each error handler label in a procedure a unique name that will not conflict with any other element in the procedure, and make sure you append a colon to the name. Within the procedure, place the  **Exit Sub** or **Exit Function** statement in front of the error handler label so that the procedure doesn't run the error-checking code if no error occurs.




```vb
Sub CausesAnError() 
    ' Direct procedure flow. 
    On Error GoTo ErrorHandler 
    ' Raise division by zero error. 
    Err.Raise 11 
    Exit Sub 
 
ErrorHandler: 
    ' Display error information. 
    MsgBox "Error number " &; Err.Number &; ": " &; Err.Description 
    ' Resume with statement following occurrence of error. 
    Resume Next 
End Sub
```

The  **Raise** method of the **Err** object generates the specified error. The **Number** property of the **Err** object returns the number corresponding to the most recent run-time error; the **Description** property returns the corresponding message text for a given error.

 **Note**  

* In versions 1.x and 2.0 of Access, you might have used the Error statement to generate the error, the Err function to return the error number, and the Error function to return a description of the error. Existing error-handling code that relies on the Error statement and the Error function will continue to work. However, it's better to use the Err object and its properties and methods when writing new code.

* Versions 1.x and 2.0 of Access returned only one error for all Automation, (formerly called OLE Automation) errors. The COM component application that generated the error now returns the same error information that you would receive if you were working in that application. You may need to rewrite existing error-handling code to handle new Automation errors properly.

* If you wish to return the descriptive string associated with a Access error or a Data Access Objects (DAO) error, but the error has not actually occurred in your code, you can use the AccessError method to return the string.

