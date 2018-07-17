---
title: Folder.SetCustomIcon Method (Outlook)
keywords: vbaol11.chm3317
f1_keywords:
- vbaol11.chm3317
ms.prod: outlook
api_name:
- Outlook.Folder.SetCustomIcon
ms.assetid: d368547b-e90c-85ec-7d5c-e48cbe8eb42e
ms.date: 06/08/2017
---


# Folder.SetCustomIcon Method (Outlook)

Sets a custom icon that is specified by  _Picture_ for the folder.


## Syntax

 _expression_ . **SetCustomIcon**( **_Picture_** )

 _expression_ A variable that represents a **[Folder](folder-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Picture_|Required| **[IPictureDisp](http://msdn.microsoft.com/en-us/library/ms680762%28VS.85%29.aspx)**|Specifies the custom icon for the folder.|

## Remarks

The  **IPictureDisp** object specified by _Picture_ must have its **Type** property equal to **PICTYPE_ICON** or **PICTYPE_BITMAP** . The icon or bitmap resource can have a maximum size of 32x32. Icons that are 16x16 or 24x24 are also supported, and Microsoft Outlook can scale up a 16x16 icon if Outlook is running in high Dots Per Inch (DPI) mode. Icons of other sizes cause **SetCustomIcon** to return an error.

You can set a custom icon for a search folder and for all folders that do not represent a default or a special folder. If you attempt to set a custom icon for a folder that belongs to one of the following groups of folders,  **SetCustomIcon** will return an error:


-  Default folders (as listed by the **[OlDefaultFolders](oldefaultfolders-enumeration-outlook.md)** enumeration)
    
- Special folders (as listed by the  **[OlSpecialFolders](olspecialfolders-enumeration-outlook.md)** enumeration)
    
- Exchange public folders
    
-  Root folder of any Exchange mailbox
    
- Hidden folders
    
You can only call  **SetCustomIcon** from code that runs in-process as Outlook. An **IPictureDisp** object cannot be marshaled across process boundaries. If you attempt to call **SetCustomIcon** from out-of-process code, an exception will occur. For more details, see[An automation server cannot pass a pointer to the picture object's IPictureDisp implementation across process boundaries](http://support.microsoft.com/kb/150034).

The custom folder icon that this method provides does not persist beyond the running Outlook session. Add-ins therefore must set the custom folder icon every time that Outlook boots.

The custom folder icon does not appear in other Exchange clients such as Outlook Web Access, nor does it appear in Outlook running on a Windows Mobile device.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code fragment in C# sets the icons for folders that appear in the  **Solutions** module. The code fragment depends on the `PictureDispConverter` class that is also illustrated below.




```C#
//Get the icons for the solution 
stdole.StdPicture rootPict = 
 PictureDispConverter.ToIPictureDisp( 
 Properties.Resources.BRIDGE) 
 as stdole.StdPicture; 
stdole.StdPicture calPict = 
 PictureDispConverter.ToIPictureDisp( 
 Properties.Resources.umbrella) 
 as stdole.StdPicture; 
stdole.StdPicture contactsPict = 
 PictureDispConverter.ToIPictureDisp( 
 Properties.Resources.group) 
 as stdole.StdPicture; 
stdole.StdPicture tasksPict = 
 PictureDispConverter.ToIPictureDisp( 
 Properties.Resources.SUN) 
 as stdole.StdPicture; 
 
//Set the icons for solution folders 
solutionRoot.SetCustomIcon(rootPict); 
solutionCalendar.SetCustomIcon(calPict); 
solutionContacts.SetCustomIcon(contactsPict); 
solutionTasks.SetCustomIcon(tasksPict);
```

The following is the definition of the  `PictureDispConverter` class.




```C#
using System; 
using System.Drawing; 
using System.Collections.Generic; 
using System.Runtime.InteropServices; 
 
public static class PictureDispConverter 
{ 
 //IPictureDisp GUID. 
 public static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID; 
 
 // Converts an Icon into an IPictureDisp. 
 public static stdole.IPictureDisp ToIPictureDisp(Icon icon) 
 { 
 PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon); 
 return PictureDispConverter.OleCreatePictureIndirect(pictIcon, ref iPictureDispGuid, true); 
 } 
 
 // Converts an image into an IPictureDisp. 
 public static stdole.IPictureDisp ToIPictureDisp(Image image) 
 { 
 Bitmap bitmap = (image is Bitmap) ? (Bitmap)image : new Bitmap(image); 
 PICTDESC.Bitmap pictBit = new PICTDESC.Bitmap(bitmap); 
 return PictureDispConverter.OleCreatePictureIndirect(pictBit, ref iPictureDispGuid, true); 
 } 
 
 
 [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, 
 PreserveSig = false)] 
 private static extern stdole.IPictureDisp OleCreatePictureIndirect( 
 [MarshalAs(UnmanagedType.AsAny)] object picdesc, ref Guid iid, bool fOwn); 
 
 private readonly static HandleCollector handleCollector = 
 new HandleCollector("Icon handles", 1000); 
 
 // WINFORMS COMMENT: 
 // PICTDESC is a union in native, so we'll just 
 // define different ones for the different types 
 // the "unused" fields are there to make it the right 
 // size, since the struct in native is as big as the biggest 
 // union. 
 private static class PICTDESC 
 { 
 //Picture Types 
 public const short PICTYPE_UNINITIALIZED = -1; 
 public const short PICTYPE_NONE = 0; 
 public const short PICTYPE_BITMAP = 1; 
 public const short PICTYPE_METAFILE = 2; 
 public const short PICTYPE_ICON = 3; 
 public const short PICTYPE_ENHMETAFILE = 4; 
 
 [StructLayout(LayoutKind.Sequential)] 
 public class Icon 
 { 
 internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Icon)); 
 internal int picType = PICTDESC.PICTYPE_ICON; 
 internal IntPtr hicon = IntPtr.Zero; 
 internal int unused1 = 0; 
 internal int unused2 = 0; 
 
 internal Icon(System.Drawing.Icon icon) 
 { 
 this.hicon = icon.ToBitmap().GetHicon(); 
 } 
 } 
 
 [StructLayout(LayoutKind.Sequential)] 
 public class Bitmap 
 { 
 internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Bitmap)); 
 internal int picType = PICTDESC.PICTYPE_BITMAP; 
 internal IntPtr hbitmap = IntPtr.Zero; 
 internal IntPtr hpal = IntPtr.Zero; 
 internal int unused = 0; 
 internal Bitmap(System.Drawing.Bitmap bitmap) 
 { 
 this.hbitmap = bitmap.GetHbitmap(); 
 } 
 } 
 } 
} 

```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

