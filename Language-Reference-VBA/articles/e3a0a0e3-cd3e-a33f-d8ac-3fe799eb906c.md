
# Invalid in Immediate window

 **Last modified:** July 28, 2015

Not all statements are permitted in the  **Immediate** window. This error has the following causes and solutions:




- A declarative statement was used. For example,  **Const**,  **Declare**,  **Def**_type_,  **Dim**,  **Function**,  **Option Base**,  **Option Explicit**,  **Option Compare**,  **Option Private**,  **Private**,  **Public**, property procedure declaration statements ( **Property Let**,  **Property Set**, and  **Property Get**),  **ReDim**,  **Static**,  **Sub**, and  **Type** are not allowed in the **Immediate** window. Remove the declarative statements from the **Immediate** window.
    
- A control flow statement was used, for example,  **Sub**,  **Function**,  **Property**,  **GoSub**,  **GoTo**,  **Return**, and  **Resume**. Remove these statements from the  **Immediate** window.
    
- There is no logical connection made between separated physical lines in the  **Immediate** window, so statements formatted as multiple physical lines, such as a block **If** statement, can't be properly executed. Such blocks can be typed on a single physical line, with each statement separated from the next by a colon ( **:**). Conversely, you can extend a single statement across physical lines in the  **Immediate** window by using the [line-continuation character](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), which is a space followed by an underscore (  **_**).
    
- You tried to execute some code in the  **Immediate** window that invalidates the current state of your program and requires you to reinitialize the program. Remove the code in question from the **Immediate** window.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
