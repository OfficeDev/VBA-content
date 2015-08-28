
# IRR Function

 **Last modified:** July 28, 2015


Returns a  [Double](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) specifying the internal rate of return for a series of periodic cash flows (payments and receipts).
 **Syntax**
 **IRR( _values_()**[,  **_guess_**] **)**
The  **IRR** function has these [named arguments](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md):


|**Part**|**Description**|
|:-----|:-----|
| **_values_()**|Required.  [Array](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) of **Double** specifying cash flow values. The array must contain at least one negative value (a payment) and one positive value (a receipt).|
| **_guess_**|Optional.  [Variant](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) specifying value you estimate will be returned by **IRR**. If omitted,  **_guess_** is 0.1 (10 percent).|
 **Remarks**
The internal rate of return is the interest rate received for an investment consisting of payments and receipts that occur at regular intervals.
The  **IRR** function uses the order of values within the array to interpret the order of payments and receipts. Be sure to enter your payment and receipt values in the correct sequence. The cash flow for each period doesn't have to be fixed, as it is for an annuity.
 **IRR** is calculated by iteration. Starting with the value of **_guess_**,  **IRR** cycles through the calculation until the result is accurate to within 0.00001 percent. If **IRR** can't find a result after 20 tries, it fails.

## Example

In this example, the  **IRR** function returns the internal rate of return for a series of 5 cash flows contained in the array `Values()`. The first array element is a negative cash flow representing business start-up costs. The remaining 4 cash flows represent positive cash flows for the subsequent 4 years.  `Guess` is the estimated internal rate of return.


```
Dim Guess, Fmt, RetRate, Msg
Static Values(5) As Double    ' Set up array.
Guess = .1    ' Guess starts at 10 percent.
Fmt = "#0.00"    ' Define percentage format.
Values(0) = -70000    ' Business start-up costs.
' Positive cash flows reflecting income for four successive years.
Values(1) = 22000 : Values(2) = 25000
Values(3) = 28000 : Values(4) = 31000
RetRate = IRR(Values(), Guess) * 100    ' Calculate internal rate.
Msg = "The internal rate of return for these five cash flows is "
Msg = Msg &amp; Format(RetRate, Fmt) &amp; " percent."
MsgBox Msg    ' Display internal return rate.


```

