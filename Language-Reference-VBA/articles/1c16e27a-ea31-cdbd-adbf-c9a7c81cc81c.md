
# String Data Type

 **Last modified:** July 28, 2015

There are two kinds of strings: variable-length and fixed-length strings.




- A variable-length string can contain up to approximately 2 billion (2^31) characters.
    
- A fixed-length string can contain 1 to approximately 64K (2^16) characters.
    
     **Note**  A  [Public](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) fixed-length string can't be used in a [class module](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).

The codes for  [String](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) characters range from 0-255. The first 128 characters (0-127) of the character set correspond to the letters and symbols on a standard U.S. keyboard. These first 128 characters are the same as those defined by the [ASCII](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) character set. The second 128 characters (128-255) represent special characters, such as letters in international alphabets, accents, currency symbols, and fractions. The [type-declaration character](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) for **String** is the dollar sign ( **$**).
