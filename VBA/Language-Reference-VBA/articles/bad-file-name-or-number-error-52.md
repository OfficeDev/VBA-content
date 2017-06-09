---
title: Bad file name or number (Error 52)
keywords: vblr6.chm1000052
f1_keywords:
- vblr6.chm1000052
ms.prod: office
ms.assetid: 9318e732-9cba-c4ec-2108-8147b34e0847
ms.date: 06/08/2017
---


# Bad file name or number (Error 52)

An error occurred trying to access the specified file. This error has the following causes and solutions:



- A statement refers to a file with a [file number](vbe-glossary.md) or file name that is:
    
    
    
      - Not specified in the  **Open** statement or was specified in an **Open** statement, but has since been closed. Specify the file name in an **Open** statement. Note that if you invoked the **Close** statement without[arguments](vbe-glossary.md), you may have inadvertently closed all currently open files, invalidating all file numbers.
    
  - Out of the range of file numbers (1 - 511). If your code is generating file numbers algorithmically, make sure the numbers are valid.
    

    
    
- There is an invalid name or number.
    
  ```
  LETTER.DOC 
My Memo.Txt 
BUDGET.92 
12345678.901 
Second Try.Rpt 

  ```


    File names must conform to operating system conventions as well as Basic file-naming conventions. In Microsoft Windows, use the following conventions for naming files and directories:
    
    
    
      - The name of a file or directory can have two parts: a name and an optional extension. The two parts are separated by a period, for example, myfile.new.
    
  - The name can contain up to 255 characters.
    
  - The name must start with either a letter or number. It can contain any uppercase or lowercase characters (file names aren't case-sensitive) except the following characters: quotation mark ( **"** ), apostrophe ( **'** ), slash ( **/** ), backslash ( **\** ), colon ( **:** ), and vertical bar ( **|** ).
    
  - The name can contain spaces.
    
  - The following names are reserved and can't be used for files or directories: CON, AUX, COM1, COM2, COM3, COM4, LPT1, LPT2, LPT3, PRN, and NUL. For example, if you try to name a file PRN in an  **Open** statement, the default printer will simply become the destination for **Print #** and **Write #** statements directed to the file number specified in the **Open** statement.
    
  - The following are examples of valid Microsoft Windows file names:
    

    
    
    
    
      - On the Macintosh, a file name can include any character except the colon ( **:** ), and can contain spaces. Null characters ( **Chr(0)** ) aren't allowed in any file names.
    

    
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

