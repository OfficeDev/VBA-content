
# Name Property (VBA Add-In Object Model)

 **Last modified:** July 28, 2015


Returns or sets a  [String](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) containing the name used in code to identify an object. For the **VBProject** object and the **VBComponent** object, read/write; for the **Property** object and the **Reference** object, read-only.
 **Remarks**
The following table describes how the  **Name** property setting applies to different objects.


|**Object**|**Result of Using Name Property Setting**|
|:-----|:-----|
| **VBProject**|Returns or sets the name of the active  [project](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).|
| **VBComponent**|Returns or sets the name of the component. An error occurs if you try to set the  **Name** property to a name already being used or an invalid name.|
| **Property**|Returns the name of the property as it appears in the  **Property Browser**. This is the value used to index the  **Properties** [collection](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md). The name can't be set.|
| **Reference**|Returns the name of the reference in code. The name can't be set.|
The default name for new objects is the type of object plus a unique integer. For example, the first new Form object is Form1, a new Form object is Form1, and the third TextBox control you create on a form is TextBox3.
An object's  **Name** property must start with a letter and can be a maximum of 40 characters. It can include numbers and underline (_) characters but can't include punctuation or spaces. [Forms](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) and [modules](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) can't have the same name as another public object such as **Clipboard**,  **Screen**, or  **App**. Although the  **Name** property setting can be a [keyword](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), property name, or the name of another object, this can create conflicts in your code.
