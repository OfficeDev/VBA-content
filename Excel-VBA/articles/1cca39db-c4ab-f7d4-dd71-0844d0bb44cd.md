
# WorksheetFunction.Replace Method (Excel)

 **Last modified:** July 28, 2015

Replaces part of a text string, based on the number of characters you specify, with a different text string.

## Syntax

 _expression_. **Replace**( **_Arg1_**,  **_Arg2_**,  **_Arg3_**,  **_Arg4_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **String**|Text in which you want to replace some characters.|
|Arg2|Required| **Double**|The position of the character in  **Arg1** that you want to replace with **Arg4**.|
|Arg3|Required| **Double**|The number of characters in  **Arg1** that you want the **Replace** method to replace with **Arg4**.|
|Arg4|Required| **String**|Text that will replace characters in  **Arg1**.|

### Return Value

A String value that represents the new string, after replacement.


## Example

This example replaces abcdef with ac-ef and notifies the user during this process.


```
Sub UseReplace() 
 
 Dim strCurrent As String 
 Dim strReplaced As String 
 
 strCurrent = "abcdef" 
 
 ' Notify user and display current string. 
 MsgBox "The current string is: " &amp; strCurrent 
 
 ' Replace "cd" with "-". 
 strReplaced = Application.WorksheetFunction.Replace _ 
 (Arg1:=strCurrent, Arg2:=3, _ 
 Arg3:=2, Arg4:="-") 
 
 ' Notify user and display replaced string. 
 MsgBox "The replaced string is: " &amp; strReplaced 
 
End Sub
```


## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
