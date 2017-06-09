---
title: Elements of Run-Time Error Handling
keywords: vbaac10.chm5186924
f1_keywords:
- vbaac10.chm5186924
ms.prod: access
ms.assetid: a0e06a1e-2709-aa51-92d0-340788a31a8a
ms.date: 06/08/2017
---


# Elements of Run-Time Error Handling

## Errors and Error Handling

When you're programming an application, you need to consider what happens when an error occurs. An error can occur in your application for one of two of reasons. First, some condition at the time the application is running makes otherwise valid code fail. For example, if your code attempts to open a table that the user has deleted, an error occurs. Second, your code may contain improper logic that prevents it from doing what you intended. For example, an error occurs if your code attempts to divide a value by zero.

If you've implemented no error handling, then Visual Basic halts execution and displays an error message when an error occurs in your code. The user of your application is likely to be confused and frustrated when this happens. You can forestall many problems by including thorough error-handling routines in your code to handle any error that may occur.

When adding error handling to a procedure, you should consider how the procedure will route execution when an error occurs. The first step in routing execution to an error handler is to enable an error handler by including some form of the  **On Error** statement within the procedure. The **On Error** statement directs execution in event of an error. If there's no **On Error** statement, Visual Basic simply halts execution and displays an error message when an error occurs.

When an error occurs in a procedure with an enabled error handler, Visual Basic doesn't display the normal error message. Instead it routes execution to an error handler, if one exists. When execution passes to an enabled error handler, that error handler becomes active. Within the active error handler, you can determine the type of error that occurred and address it in the manner that you choose. Access provides three objects that contain information about errors that have occurred, the ADO  **Error** object, the Visual Basic[Err Object](http://msdn.microsoft.com/library/23c9697a-9c6b-18f8-2b86-a0735f082c67%28Office.15%29.aspx) **Err** object, and the DAO **Error** object.


## Routing Execution When an Error Occurs

An error handler specifies what happens within a procedure when an error occurs. For example, you may want the procedure to end if a certain error occurs, or you may want to correct the condition that caused the error and resume execution. The  **On Error** and **Resume** statements determine how execution proceeds in the event of an error.

 _The On Error Statement_

The  **On Error** statement enables or disables an error-handling routine. If an error-handling routine is enabled, execution passes to the error-handling routine when an error occurs.

There are three forms of the  **On Error** statement: **On Error GoTo** _label_, **On Error GoTo 0**, and **On Error Resume Next**. The **On Error GoTo** _label_ statement enables an error-handling routine, beginning with the line on which the statement is found. You should enable the error-handling routine before the first line at which an error could occur. When the error handler is active and an error occurs, execution passes to the line specified by the _label_ argument.

The line specified by the  _label_ argument should be the beginning of the error-handling routine. For example, the following procedure specifies that if an error occurs, execution passes to the line labeled :




```vb
Function MayCauseAnError() 
    ' Enable error handler. 
    On Error GoTo Error_MayCauseAnError 
    .            ' Include code here that may generate error. 
    . 
    . 
 
Error_MayCauseAnError: 
    .            ' Include code here to handle error. 
    . 
    . 
End Function
```

The  **On Error GoTo 0** statement disables error handling within a procedure. It doesn't specify line 0 as the start of the error-handling code, even if the procedure contains a line numbered 0. If there's no **On Error GoTo 0** statement in your code, the error handler is automatically disabled when the procedure has run completely. The **On Error GoTo 0** statement resets the properties of the **Err** object, having the same effect as the **Clear** method of the **Err** object.

The  **On Error Resume Next** statement ignores the line that causes an error and routes execution to the line following the line that caused the error. Execution isn't interrupted. You can use the **On Error Resume Next** statement if you want to check the properties of the **Err** object immediately after a line at which you anticipate an error will occur, and handle the error within the procedure rather than in an error handler.

 _The Resume Statement_

The  **Resume** statement directs execution back to the body of the procedure from within an error-handling routine. You can include a **Resume** statement within an error-handling routine if you want execution to continue at a particular point in a procedure. However, a **Resume** statement isn't necessary; you can also end the procedure after the error-handling routine.

There are three forms of the  **Resume** statement. The **Resume** or **Resume 0** statement returns execution to the line at which the error occurred. The **Resume Next** statement returns execution to the line immediately following the line at which the error occurred. The **Resume** _label_ statement returns execution to the line specified by the _label_ argument. The _label_ argument must indicate either a line label or a line number.

You typically use the  **Resume** or **Resume 0** statement when the user must make a correction. For example, if you prompt the user for the name of a table to open, and the user enters the name of a table that doesn't exist, you can prompt the user again and resume execution with the statement that caused the error.

You use the  **Resume Next** statement when your code corrects for the error within an error handler, and you want to continue execution without rerunning the line that caused the error. You use the **Resume** _label_ statement when you want to continue execution at another point in the procedure, specified by the _label_ argument. For example, you might want to resume execution at an exit routine, as described in the following section.

 _Exiting a Procedure_

When you include an error-handling routine in a procedure, you should also include an exit routine, so that the error-handling routine will run only if an error occurs. You can specify an exit routine with a line label in the same way that you specify an error-handling routine.

For example, you can add an exit routine to the example in the previous section. If an error doesn't occur, the exit routine runs after the body of the procedure. If an error occurs, then execution passes to the exit routine after the code in the error-handling routine has run. The exit routine contains an  **Exit** statement.




```vb
Function MayCauseAnError() 
    ' Enable error handler. 
    On Error GoTo Error_MayCauseAnError 
    .            ' Include code here that may generate error. 
    . 
    . 
 
Exit_MayCauseAnError: 
    Exit Function 
 
Error_MayCauseAnError: 
    .            ' Include code to handle error. 
    . 
    . 
    ' Resume execution with exit routine to exit function. 
    Resume Exit_MayCauseAnError 
End Function
```

 _Handling Errors in Nested Procedures_

When an error occurs in a nested procedure that doesn't have an enabled error handler, Visual Basic searches backward through the calls list for an enabled error handler in another procedure, rather than simply halting execution. This provides your code with an opportunity to correct the error within another procedure. For example, suppose Procedure A calls Procedure B, and Procedure B calls Procedure C. If an error occurs in Procedure C and there's no enabled error handler, Visual Basic checks Procedure B, then Procedure A, for an enabled error handler. If one exists, execution passes to that error handler. If not, execution halts and an error message is displayed.

Visual Basic also searches backward through the calls list for an enabled error handler when an error occurs within an active error handler. You can force Visual Basic to search backward through the calls list by raising an error within an active error handler with the  **Raise** method of the **Err** object. This is useful for handling errors that you don't anticipate within an error handler. If an unanticipated error occurs, and you regenerate that error within the error handler, then execution passes back up the calls list to find another error handler, which may be set up to handle the error.

For example, suppose Procedure C has an enabled error handler, but the error handler doesn't correct for the error that has occurred. Once the error handler has checked for all the errors that you've anticipated, it can regenerate the original error. Execution then passes back up the calls list to the error handler in Procedure B, if one exists, providing an opportunity for this error handler to correct the error. If no error handler exists in Procedure B, or if it fails to correct for the error and regenerates it again, then execution passes to the error handler in Procedure A, assuming one exists.

To illustrate this concept in another way, suppose that you have a nested procedure that includes error handling for a type mismatch error, an error which you've anticipated. At some point, a division-by-zero error, which you haven't anticipated, occurs within Procedure C. If you've included a statement to regenerate the original error, then execution passes back up the calls list to another enabled error handler, if one exists. If you've corrected for a division-by-zero error in another procedure in the calls list, then the error will be corrected. If your code doesn't regenerate the error, then the procedure continues to run without correcting the division-by-zero error. This in turn may cause other errors within the set of nested procedures.

In summary, Visual Basic searches back up the calls list for an enabled error handler if:


- An error occurs in a procedure that doesn't include an enabled error handler.
    
- An error occurs within an active error handler. If you use the  **Raise** method of the **Err** object to raise an error, you can force Visual Basic to search backward through the calls list for an enabled error handler.
    

## Getting Information About an Error

Once execution has passed to the error-handling routine, your code must determine which error has occurred and address it. Visual Basic and Access provide several language elements that you can use to get information about a specific error. Each is suited to different types of errors. Since errors can occur in different parts of your application, you need to determine which to use in your code based on what errors you expect.

The language elements available for error handling include:


- The  **Err** object.
    
- The ADO  **Error** object and **Errors** collection
    
- The DAO  **Error** object and **Errors** collection.
    
- The  **AccessError** method.
    
- The  **Error** event.
    
 _The Err Object_

The  **Err** object is provided by Visual Basic. When a Visual Basic error occurs, information about that error is stored in the **Err** object. The **Err** object maintains information about only one error at a time. When a new error occurs, the **Err** object is updated to include information about that error instead.

To get information about a particular error, you can use the properties and methods of the  **Err** object. The **Number** property is the default property of the **Err** object; it returns the identifying number of the error that occurred. The **Err** object's **Description** property returns the descriptive string associated with a Visual Basic error. The **Clear** method clears the current error information from the **Err** object. The **Raise** method generates a specific error and populates the properties of the **Err** object with information about that error.

The following example shows how to use the  **Err** object in a procedure that may cause a type mismatch error:




```vb
Function MayCauseAnError() 
    ' Declare constant to represent likely error. 
    Const conTypeMismatch As Integer = 13 
 
    On Error GoTo Error_MayCauseAnError 
        .            ' Include code here that may generate error. 
        . 
        . 
 
Exit_MayCauseAnError: 
    Exit Function 
 
Error_MayCauseAnError: 
    ' Check Err object properties. 
    If Err = conTypeMismatch Then 
        .            ' Include code to handle error. 
        . 
        . 
    Else 
        ' Regenerate original error. 
        Dim intErrNum As Integer 
        intErrNum = Err 
        Err.Clear 
        Err.Raise intErrNum 
    End If 
    ' Resume execution with exit routine to exit function. 
    Resume Exit_MayCauseAnError 
End Function
```

Note that in the preceding example, the  **Raise** method is used to regenerate the original error. If an error other than a type mismatch error occurs, execution will be passed back up the calls list to another enabled error handler, if one exists.

The  **Err** object provides you with all the information you need about Visual Basic errors. However, it doesn't give you complete information about Access errors or Access database engine errors. Access and Data Access Objects (DAO)) provide additional language elements to assist you with those errors.

 _The Error Object and Errors Collection_

The  **Error** object and **Errors** collection are provided by ADO and DAO. The **Error** object represents an ADO or DAO error. A single ADO or DAO operation may cause several errors, especially if you are performing DAO ODBC operations. Each error that occurs during a particular data access operation has an associated **Error** object. All the **Error** objects associated with a particular ADO or DAO operation are stored in the **Errors** collection, the lowest-level error being the first object in the collection and the highest-level error being the last object in the collection.

When a ADO or DAO error occurs, the Visual Basic  **Err** object contains the error number for the first object in the **Errors** collection. To determine whether additional ADO or DAO errors have occurred, check the **Errors** collection. The values of the ADO **Number** or DAO **Number** properties and the ADO **Description** or DAO **Description** properties of the first **Error** object in the **Errors** collection should match the values of the **Number** and **Description** properties of the Visual Basic **Err** object.

 _The AccessError Method_

You can use the  **Raise** method of the **Err** object to generate a Visual Basic error that hasn't actually occurred and determine the descriptive string associated with that error. However, you can't use the **Raise** method to generate a Access error, an ADO error, or a DAO error. To determine the descriptive string associated with an Access error, an ADO error, or a DAO error that hasn't actually occurred, use the **AccessError** method.

 _The Error Event_

You can use the Error event to trap errors that occur on an Access form or report. For example, if a user tries to enter text in a field whose data type is Date/Time, the Error event occurs. If you add an Error event procedure to an Employees form, then try to enter a text value in the HireDate field, the Error event procedure runs.

The Error event procedure takes an integer argument, DataErr. When an Error event procedure runs, the DataErr argument contains the number of the Access error that occurred. Checking the value of the DataErr argument within the event procedure is the only way to determine the number of the error that occurred. The  **Err** object isn't populated with error information after the Error event occurs. You can use the value of the DataErr argument together with the **AccessError** method to determine the number of the error and its descriptive string.


 **Note**  The  **Error** statement and **Error** function are provided for backward compatibility only. When writing new code, use the **Err** and **Error** objects, the **AccessError** function, and the Error event for getting information about an error.

 **Link provided by:** The[UtterAccess](http://www.utteraccess.com) community


- [Handling Access Errors with VBA](http://www.utteraccess.com/wiki/index.php/Error_Handling)
    

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


