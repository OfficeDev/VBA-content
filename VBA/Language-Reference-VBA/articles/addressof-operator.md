---
title: AddressOf Operator
keywords: vblr6.chm1103681
f1_keywords:
- vblr6.chm1103681
ms.prod: office
ms.assetid: 2c966af2-ad6d-5eab-9202-0a2b7780baf4
ms.date: 06/08/2017
---


# AddressOf Operator



A unary operator that causes the address of the [procedure](vbe-glossary.md) it precedes to be passed to an API procedure that expects a function pointer at that position in the[argument](vbe-glossary.md) list.
 **Syntax**
 **AddressOf**_procedurename_
The required  _procedurename_ specifies the procedure whose address is to be passed. It must represent a procedure in a[standard module](vbe-glossary.md) module in the[project](vbe-glossary.md) in which the call is made.
 **Remarks**
When a procedure name appears in an argument list, usually the procedure is evaluated, and the address of the procedure's return value is passed.  **AddressOf** permits the address of the procedure to be passed to a Windows API function in a[dynamic-link library (DLL)](vbe-glossary.md), rather passing the procedure's return value. The API function can then use the address to call the Basic procedure, a process known as a callback. The  **AddressOf** operator appears only in the call to the API procedure.
Although you can use  **AddressOf** to pass procedure pointers among Basic procedures, you can't call a function through such a pointer from within Basic. This means, for example, that a[class](vbe-glossary.md) written in Basic can't make a callback to its controller using such a pointer. When using **AddressOf** to pass a procedure pointer among procedures within Basic, the[parameter](vbe-glossary.md) of the called procedure must be typed **As Long**.
Using  **AddressOf** may cause unpredictable results if you don't completely understand the concept of function callbacks. You must understand how the Basic portion of the callback works, and also the code of the DLL into which you are passing your function address. Debugging such interactions is difficult since the program runs in the same process as the[development environment](vbe-glossary.md). In some cases, systematic debugging may not be possible.

 **Note**  You can create your own call-back function prototypes in DLLs compiled with Microsoft Visual C++ (or similar tools). To work with  **AddressOf**, your prototype must use the __stdcall calling convention. The default calling convention (__cdecl) will not work with **AddressOf**.

Since the caller of a callback is not within your program, it is important that an error in the callback procedure not be propagated back to the caller. You can accomplish this by placing the  **On Error Resume Next** statement at the beginning of the callback procedure.

## Example

The following example creates a form with a list box containing an alphabetically sorted list of the fonts in your system.

To run this example, create a form with a list box on it. The code for the form is as follows:




```vb
Option Explicit

Private Sub Form_Load()
    Module1.FillListWithFonts List1
End Sub
```

Place the following code in a module. The third argument in the definition of the EnumFontFamilies function is a  **Long** that represents a procedure. The argument must contain the address of the procedure, rather than the value that the procedure returns. In the call to EnumFontFamilies, the third argument requires the **AddressOf** operator to return the address of the EnumFontFamProc procedure, which is the name of the callback procedure you supply when calling the Windows API function, **EnumFontFamilies**. Windows calls EnumFontFamProc once for each of the font families on the system when you pass **AddressOf** EnumFontFamProc to **EnumFontFamilies**. The last argument passed to **EnumFontFamilies** specifies the list box in which the information is displayed.




```vb
'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &;H40&;
Public Const NTM_BOLD = &;H20&;
Public Const NTM_ITALIC = &;H1&;

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &;H1
Public Const TMPF_VECTOR = &;H2
Public Const TMPF_DEVICE = &;H8
Public Const TMPF_TRUETYPE = &;H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &;H1
Public Const DEVICE_FONTTYPE = &;H2
Public Const TRUETYPE_FONTTYPE = &;H4

Declare Function EnumFontFamilies Lib "gdi32" Alias _
     "EnumFontFamiliesA" _
     (ByVal hDC As Long, ByVal lpszFamily As String, _ 
     ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
     ByVal hDC As Long) As Long

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _ 
     ByVal FontType As Long, LParam As ListBox) As Long
Dim FaceName As String
Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    LParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function

Sub FillListWithFonts(LB As ListBox)
Dim hDC As Long
    LB.Clear
    hDC = GetDC(LB.hWnd)
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, LB
    ReleaseDC LB.hWnd, hDC
End Sub
```


