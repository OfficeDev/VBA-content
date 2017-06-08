---
title: TextBox.InputMask Property (Access)
keywords: vbaac10.chm11046
f1_keywords:
- vbaac10.chm11046
ms.prod: access
api_name:
- Access.TextBox.InputMask
ms.assetid: a705c2a4-ff2f-74d1-4a7c-1eade3b00ae8
ms.date: 06/08/2017
---


# TextBox.InputMask Property (Access)

You can use the  **InputMask** property to make data entry easier and to control the values users can enter in a text boxcontrol. Read/write **String**.


## Syntax

 _expression_. **InputMask**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Input masks are helpful for data-entry operations such as an input mask for a Phone Number field that shows you exactly how to enter a new number: (___) ___-____. It is often easier to use the Input Mask Wizard to set the property for you.

The  **InputMask** property can contain up to three sections separated by semicolons (;).



|**Section**|**Description**|
|:-----|:-----|
|First|Specifies the input mask itself; for example, !(999) 999-9999. For a list of characters you can use to define the input mask, see the following table.|
|Second|Specifies whether Microsoft Access stores the literal display characters in the table when you enter data. If you use 0 for this section, all literal display characters (for example, the parentheses in a phone number input mask) are stored with the value; if you enter 1 or leave this section blank, only characters typed into the control are stored.|
|Third|Specifies the character that Microsoft Access displays for the space where you should type a character in the input mask. For this section, you can use any character; to display an empty string, use a space enclosed in quotation marks (" ").|
In Visual Basic you use a string expression to set this property. For example, the following specifies an input mask for a text box control used for entering a phone number:




```vb
Forms!Customers!Telephone.InputMask = "(###) ###-####"
```

When you create an input mask, you can use special characters to require that certain data be entered (for example, the area code for a phone number) and that other data be optional (such as a telephone extension). These characters specify the type of data, such as a number or character, that you must enter for each character in the input mask.

You can define an input mask by using the following characters.



|**Character**|**Description**|
|:-----|:-----|
|0|Digit (0 to 9, entry required, plus [+] and minus [?] signs not allowed).|
|9|Digit or space (entry not required, plus and minus signs not allowed).|
|#|Digit or space (entry not required; spaces are displayed as blanks while in Edit mode, but blanks are removed when data is saved; plus and minus signs allowed).|
|L|Letter (A to Z, entry required).|
|?|Letter (A to Z, entry optional).|
|A|Letter or digit (entry required).|
|a|Letter or digit (entry optional).|
|&;|Any character or a space (entry required).|
|C|Any character or a space (entry optional).|
|. , : ; - /|Decimal placeholder and thousand, date, and time separators. (The actual character used depends on the settings in the  **Regional Settings Properties** dialog box in Windows Control Panel).|
|<|Causes all characters to be converted to lowercase.|
|>|Causes all characters to be converted to uppercase.|
|!|Causes the input mask to display from right to left, rather than from left to right. Characters typed into the mask always fill it from left to right. You can include the exclamation point anywhere in the input mask.|
|\|Causes the character that follows to be displayed as the literal character (for example, \A is displayed as just A).|

 **Note**   Setting the **InputMask** property to the word "Password" creates a password-entry control. Any character typed in the control is stored as the character but is displayed as an asterisk (*). You use the Password input mask to prevent displaying the typed characters on the screen.

When you type data in a field for which you've defined an input mask, the data is always entered in Overtype mode. If you use the BACKSPACE key to delete a character, the character is replaced by a blank space.

If you move text from a field for which you've defined an input mask onto the Clipboard, the literal display characters are copied, even if you have specified that they not be saved with data.


 **Note**  Only characters that you type directly in a control or combo box are affected by the input mask. Microsoft Access ignores any input masks when you import data, run an action query, or enter characters in a control by setting the control's  **Text** property in Visual Basic or by using the SetValue action in a macro.

When you've defined an input mask and set the  **Format** property for the same field, the **Format** property takes precedence when the data is displayed. This means that even if you've saved an input mask, the input mask is ignored when data is formatted and displayed. The data in the underlying table itself isn't changed; the **Format** property affects only how the data is displayed.

The following table shows some useful input masks and the type of values you can enter in them.



|**Input mask**|**Sample values**|
|:-----|:-----|
|(000) 000-0000|(206) 555-0248|
|(999) 999-9999|(206) 555-0248|
||( ) 555-0248|
|(000) AAA-AAAA|(206) 555-TELE|
|#999|?20|
||2000|
|>L????L?000L0|GREENGR339M3|
||MAY R 452B7|
|>L0L 0L0|T2F 8M4|
|00000-9999|98115-|
||98115-3007|
|>L<??????????????|Maria|
||Brendan|
|SSN 000-00-0000|SSN 555-55-5555|
|>LL00000-0000|DB51392-0493|

## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

