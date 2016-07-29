
# SetEOS Method (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Sets the position that is the end of the stream.

## Syntax

 _Stream_. **SetEOS**


## Remarks

 **SetEOS** updates the value of the[EOS](97cd23ef-cca8-4dcc-2641-082a0e1b853c.md) property, by making the current[Position](a07c9197-673b-ddf2-fca9-b0b54fbd67b4.md) the end of the stream. Any bytes or characters following the current position are truncated.

Since [Write](cabe4581-409f-7f05-bd59-d495bfb2c6fd.md), [WriteText](1ca2d9d5-11f4-d088-6fc3-53240208bb09.md), and [CopyTo](1c1ab950-51f7-7ecc-ccd8-e689db02f06a.md) do not truncate any extra values in existing **Stream** objects, you can truncate these bytes or characters by setting the new end-of-stream position with **SetEOS**.


 **Caution**  If you set  **EOS** to a position before the actual end of the stream, you will lose all data after the new **EOS** position.

