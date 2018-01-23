---
title: Dir Function
keywords: vblr6.chm1008898
f1_keywords:
- vblr6.chm1008898
ms.prod: office
ms.assetid: eaf6fe6e-342a-5038-3914-bb5e58fcad5a
ms.date: 06/08/2017
---


# Dir Function



Returns a  **String** representing the name of a file, directory, or folder that matches a specified pattern or file attribute, or the volume label of a drive.
 **Syntax**
 **Dir** [ **(**_pathname_ [ **,**_attributes_ ] **)** ]
The  **Dir** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _pathname_|Optional. [String expression](vbe-glossary.md) that specifies a file name — may include directory or folder, and drive. A zero-length string ("") is returned if _pathname_ is not found.|
| _attributes_|Optional. [Constant](vbe-glossary.md) or[numeric expression](vbe-glossary.md), whose sum specifies file attributes. If omitted, returns files that match  _pathname_ but have no attributes.|
 **Settings**
The  _attributes_[argument](vbe-glossary.md) settings are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbNormal**|0|(Default) Specifies files with no attributes.|
|**vbReadOnly**|1|Specifies read-only files in addition to files with no attributes.|
|**vbHidden**|2|Specifies hidden files in addition to files with no attributes.|
|**VbSystem**|4|Specifies system files in addition to files with no attributes. Not available on the Macintosh.|
|**vbVolume**|8|Specifies volume label; if any other attributed is specified,  **vbVolume** is ignored. Not available on the Macintosh.|
|**vbDirectory**|16|Specifies directories or folders in addition to files with no attributes.|
|**vbAlias**|64|Specified file name is an alias. Available only on the Macintosh.|

 **Note**  These constants are specified by Visual Basic for Applications and can be used anywhere in your code in place of the actual values.

 **Remarks**
In Microsoft Windows,  **Dir** supports the use of multiple character ( **\*** ) and single character ( **?** ) wildcards to specify multiple files. On the Macintosh, these characters are treated as valid file name characters and can't be used as wildcards to specify multiple files.
Since the Macintosh doesn't support the wildcards, use the file type to identify groups of files. You can use the  **MacID** function to specify file type instead of using the file names. For example, the following statement returns the name of the first TEXT file in the current folder:



```
Dir("SomePath", MacID("TEXT"))


```

To iterate over all files in a folder, specify an empty string:



```
Dir("")

```

If you use the  **MacID** function with **Dir** in Microsoft Windows, an error occurs.
Any  _attribute_ value greater than 256 is considered a **MacID** value.
You must specify  _pathname_ the first time you call the **Dir** function, or an error occurs. If you also specify file attributes, _pathname_ must be included.
 **Dir** returns the first file name that matches _pathname_. To get any additional file names that match _pathname_, call **Dir** again with no arguments. When no more file names match, **Dir** returns a zero-length string (""). Once a zero-length string is returned, you must specify _pathname_ in subsequent calls or an error occurs. You can change to a new _pathname_ without retrieving all of the file names that match the current _pathname_. However, you can't call the **Dir** function recursively. Calling **Dir** with the **vbDirectory** attribute does not continually return subdirectories.
With Excel for Mac 2016, the initial **Dir** function call will succeed. Subsequent calls to iterate through the specified directory will cause an error however. This is a known bug unfortunately.

 **Tip**  Because file names are retrieved in no particular order, you may want to store returned file names in an [array](vbe-glossary.md), and then sort the array.


