
# Application.DisplayAlerts Property (Excel)

 **True** if Microsoft Excel displays certain alerts and messages while a macro is running. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayAlerts**

 _expression_ A variable that represents an **Application** object.


## Remarks

The default value is  **True** . Set this property to **False** to suppress prompts and alert messages while a macro is running; when a message requires a response, Microsoft Excel chooses the default response.

If you set this property to  **False** , Microsoft Excel sets this property to **True** when the code is finished, unless you are running cross-process code.




 **Note**  When using the  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** method for workbooks to overwrite an existing file, the **Confirm Save As** dialog box has a default of **No**, while the  **Yes** response is selected by Excel when the **DisplayAlerts** property is set to **False** . The **Yes** response overwrites the existing file.When using the  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** method for workbooks to save a workbook that contains a Visual Basic for Applications (VBA) project in the Excel 5.0/95 file format, the **Microsoft Excel** dialog box has a default of **Yes**, while the  **Cancel** response is selected by Excel when the **DisplayAlerts** property is set to **False** . You cannot save a workbook that contains a VBA project using the Excel 5.0/95 file format.


## Example

This example closes the Workbook Book1.xls and does not prompt the user to save changes. Changes to Book1.xls are not saved.


```vb
Application.DisplayAlerts = False 
Workbooks("BOOK1.XLS").Close 
Application.DisplayAlerts = True
```

This example suppresses the message that otherwise appears when you initiate a DDE channel to an application that is not running.




```vb
Application.DisplayAlerts = False 
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DisplayAlerts = True 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber 
Application.DisplayAlerts = True
```


## See also


#### Concepts


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Other resources


[Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
