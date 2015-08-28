
# MultiLine Property

 **Last modified:** July 28, 2015


Specifies whether a control can accept and display multiple lines of text.
 **Syntax**
 _object_. **MultiLine** [= _Boolean_]
The  **MultiLine** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control supports more than one line of text.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
| **True**|The text is displayed across multiple lines (default).|
| **False**|The text is not displayed across multiple lines.|
 **Remarks**
A multiline  **TextBox** allows absolute line breaks and adjusts its quantity of lines to accommodate the amount of text it holds. If needed, a multiline control can have vertical scroll bars.
A single-line  **TextBox** doesn't allow absolute line breaks and doesn't use vertical scroll bars.
Single-line controls ignore the value of the  **WordWrap** property.

 **Note**  If you change  **MultiLine** to **False** in a multiline **TextBox**, all the characters in the  **TextBox** will be combined into one line, including non-printing characters (such as carriage returns and new-lines).

