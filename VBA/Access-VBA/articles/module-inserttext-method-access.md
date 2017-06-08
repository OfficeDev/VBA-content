---
title: Module.InsertText Method (Access)
keywords: vbaac10.chm12271
f1_keywords:
- vbaac10.chm12271
ms.prod: access
api_name:
- Access.Module.InsertText
ms.assetid: 105c77fe-29a3-ef93-3d01-8420f7725325
ms.date: 06/08/2017
---


# Module.InsertText Method (Access)

The  **InsertText** method inserts a specified string of text into a standard module or a class module.


## Syntax

 _expression_. **InsertText**( ** _Text_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required|**String**|The text to be inserted into the module.|

### Return Value

Nothing


## Remarks

When you insert a string by using the  **InsertText** method, Microsoft Access places the new text at the end of the module, after all other procedures.

To add multiple lines, include the intrinsic constant **vbCrLf** at the desired line breaks within the string that makes up the _text_ argument. This constant forces a carriage return and line feed.

To specify at which line the text is inserted, use the  **[InsertLines](module-insertlines-method-access.md)** method. To insert code into the Declarations section of the module, use the **InsertLines** method rather than the **InsertText** method.




 **Note**  In previous versions of Microsoft Access, the  **InsertText** method was a method of the **[Application](application-object-access.md)** object. You can still use the **InsertText** method of the **Application** object, but it's recommended that you use the **InsertText** method of the **Module** object instead.


## Example

The following example inserts a string of text into a standard module:


```vb
Function InsertProc(strModuleName) As Boolean 
 Dim mdl As Module, strText As String 
 
 On Error GoTo Error_InsertProc 
 ' Open module. 
 DoCmd.OpenModule strModuleName 
 ' Return reference to Module object. 
 Set mdl = Modules(strModuleName) 
 ' Initialize string variable. 
 strText = "Sub DisplayMessage()" &; vbCrLf _ 
 &; vbTab &; "MsgBox ""Wild!""" &; vbCrLf _ 
 &; "End Sub" 
 ' Insert text into module. 
 mdl.InsertText strText 
 InsertProc = True 
 
Exit_InsertProc: 
 Exit Function 
 
Error_InsertProc: 
 MsgBox Err &; ": " &; Err.Description 
 InsertProc = False 
 Resume Exit_InsertProc 
End Function
```


## See also


#### Concepts


[Module Object](module-object-access.md)

