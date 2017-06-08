---
title: WorksheetFunction Object (Excel)
keywords: vbaxl10.chm136072
f1_keywords:
- vbaxl10.chm136072
ms.prod: excel
api_name:
- Excel.WorksheetFunction
ms.assetid: 7b1d5639-363d-632c-2cf0-2232562646b6
ms.date: 06/08/2017
---


# WorksheetFunction Object (Excel)

Used as a container for Microsoft Excel worksheet functions that can be called from Visual Basic.


## Example

Use the  **[WorksheetFunction](application-worksheetfunction-property-excel.md)** property to return the **WorksheetFunction** object. The following example displays the result of applying the **Min** worksheet function to the range A1:C10.


```
Set myRange = Worksheets("Sheet1").Range("A1:C10") 
answer = Application.WorksheetFunction.Min(myRange) 
MsgBox answer
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example uses the  **CountA** worksheet function to determine how many cells in column A contain a value. For this example, the values in column A should be text. This example does a spell check on each value in column A, and if the value is spelled incorrectly, inserts the text "Wrong" into column B; otherwise, it inserts the text "OK" into column B.




```
Sub StartSpelling()
   'Set up your variables
   Dim iRow As Integer
   
   'And define your error handling routine.
   On Error GoTo ERRORHANDLER
   
   'Go through all the cells in column A, and perform a spellcheck on the value.
   'If the value is spelled incorrectly, write "Wrong" in column B, otherwise write "OK".
   For iRow = 1 To WorksheetFunction.CountA(Columns(1))
      If Application.CheckSpelling( _
         Cells(iRow, 1).Value, , True) = False Then
         Cells(iRow, 2).Value = "Wrong"
      Else
         Cells(iRow, 2).Value = "OK"
      End If
   Next iRow
   Exit Sub

    'Error handling routine.
ERRORHANDLER:
    MsgBox "The spell check feature is not installed!"
    
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## Methods
<a name="AboutContributor"> </a>



|**Name**|
|:-----|
|[AccrInt](worksheetfunction-accrint-method-excel.md)|
|[AccrIntM](worksheetfunction-accrintm-method-excel.md)|
|[Acos](worksheetfunction-acos-method-excel.md)|
|[Acosh](worksheetfunction-acosh-method-excel.md)|
|[Acot](worksheetfunction-acot-method-excel.md)|
|[Acoth](worksheetfunction-acoth-method-excel.md)|
|[Aggregate](worksheetfunction-aggregate-method-excel.md)|
|[AmorDegrc](worksheetfunction-amordegrc-method-excel.md)|
|[AmorLinc](worksheetfunction-amorlinc-method-excel.md)|
|[And](worksheetfunction-and-method-excel.md)|
|[Arabic](worksheetfunction-arabic-method-excel.md)|
|[Asc](worksheetfunction-asc-method-excel.md)|
|[Asin](worksheetfunction-asin-method-excel.md)|
|[Asinh](worksheetfunction-asinh-method-excel.md)|
|[Atan2](worksheetfunction-atan2-method-excel.md)|
|[Atanh](worksheetfunction-atanh-method-excel.md)|
|[AveDev](worksheetfunction-avedev-method-excel.md)|
|[Average](worksheetfunction-average-method-excel.md)|
|[AverageIf](worksheetfunction-averageif-method-excel.md)|
|[AverageIfs](worksheetfunction-averageifs-method-excel.md)|
|[BahtText](worksheetfunction-bahttext-method-excel.md)|
|[Base](worksheetfunction-base-method-excel.md)|
|[BesselI](worksheetfunction-besseli-method-excel.md)|
|[BesselJ](worksheetfunction-besselj-method-excel.md)|
|[BesselK](worksheetfunction-besselk-method-excel.md)|
|[BesselY](worksheetfunction-bessely-method-excel.md)|
|[Beta_Dist](worksheetfunction-beta_dist-method-excel.md)|
|[Beta_Inv](worksheetfunction-beta_inv-method-excel.md)|
|[BetaDist](worksheetfunction-betadist-method-excel.md)|
|[BetaInv](worksheetfunction-betainv-method-excel.md)|
|[Bin2Dec](worksheetfunction-bin2dec-method-excel.md)|
|[Bin2Hex](worksheetfunction-bin2hex-method-excel.md)|
|[Bin2Oct](worksheetfunction-bin2oct-method-excel.md)|
|[Binom_Dist](worksheetfunction-binom_dist-method-excel.md)|
|[Binom_Dist_Range](worksheetfunction-binom_dist_range-method-excel.md)|
|[Binom_Inv](worksheetfunction-binom_inv-method-excel.md)|
|[BinomDist](worksheetfunction-binomdist-method-excel.md)|
|[Bitand](worksheetfunction-bitand-method-excel.md)|
|[Bitlshift](worksheetfunction-bitlshift-method-excel.md)|
|[Bitor](worksheetfunction-bitor-method-excel.md)|
|[Bitrshift](worksheetfunction-bitrshift-method-excel.md)|
|[Bitxor](worksheetfunction-bitxor-method-excel.md)|
|[Ceiling](worksheetfunction-ceiling-method-excel.md)|
|[Ceiling_Math](worksheetfunction-ceiling_math-method-excel.md)|
|[Ceiling_Precise](worksheetfunction-ceiling_precise-method-excel.md)|
|[ChiDist](worksheetfunction-chidist-method-excel.md)|
|[ChiInv](worksheetfunction-chiinv-method-excel.md)|
|[ChiSq_Dist](worksheetfunction-chisq_dist-method-excel.md)|
|[ChiSq_Dist_RT](worksheetfunction-chisq_dist_rt-method-excel.md)|
|[ChiSq_Inv](worksheetfunction-chisq_inv-method-excel.md)|
|[ChiSq_Inv_RT](worksheetfunction-chisq_inv_rt-method-excel.md)|
|[ChiSq_Test](worksheetfunction-chisq_test-method-excel.md)|
|[ChiTest](worksheetfunction-chitest-method-excel.md)|
|[Choose](worksheetfunction-choose-method-excel.md)|
|[Clean](worksheetfunction-clean-method-excel.md)|
|[Combin](worksheetfunction-combin-method-excel.md)|
|[Combina](worksheetfunction-combina-method-excel.md)|
|[Complex](worksheetfunction-complex-method-excel.md)|
|[Confidence](worksheetfunction-confidence-method-excel.md)|
|[Confidence_Norm](worksheetfunction-confidence_norm-method-excel.md)|
|[Confidence_T](worksheetfunction-confidence_t-method-excel.md)|
|[Convert](worksheetfunction-convert-method-excel.md)|
|[Correl](worksheetfunction-correl-method-excel.md)|
|[Cosh](worksheetfunction-cosh-method-excel.md)|
|[Cot](worksheetfunction-cot-method-excel.md)|
|[Coth](worksheetfunction-coth-method-excel.md)|
|[Count](worksheetfunction-count-method-excel.md)|
|[CountA](worksheetfunction-counta-method-excel.md)|
|[CountBlank](worksheetfunction-countblank-method-excel.md)|
|[CountIf](worksheetfunction-countif-method-excel.md)|
|[CountIfs](worksheetfunction-countifs-method-excel.md)|
|[CoupDayBs](worksheetfunction-coupdaybs-method-excel.md)|
|[CoupDays](worksheetfunction-coupdays-method-excel.md)|
|[CoupDaysNc](worksheetfunction-coupdaysnc-method-excel.md)|
|[CoupNcd](worksheetfunction-coupncd-method-excel.md)|
|[CoupNum](worksheetfunction-coupnum-method-excel.md)|
|[CoupPcd](worksheetfunction-couppcd-method-excel.md)|
|[Covar](worksheetfunction-covar-method-excel.md)|
|[Covariance_P](worksheetfunction-covariance_p-method-excel.md)|
|[Covariance_S](worksheetfunction-covariance_s-method-excel.md)|
|[CritBinom](worksheetfunction-critbinom-method-excel.md)|
|[Csc](worksheetfunction-csc-method-excel.md)|
|[Csch](worksheetfunction-csch-method-excel.md)|
|[CumIPmt](worksheetfunction-cumipmt-method-excel.md)|
|[CumPrinc](worksheetfunction-cumprinc-method-excel.md)|
|[DAverage](worksheetfunction-daverage-method-excel.md)|
|[Days](worksheetfunction-days-method-excel.md)|
|[Days360](worksheetfunction-days360-method-excel.md)|
|[Db](worksheetfunction-db-method-excel.md)|
|[Dbcs](worksheetfunction-dbcs-method-excel.md)|
|[DCount](worksheetfunction-dcount-method-excel.md)|
|[DCountA](worksheetfunction-dcounta-method-excel.md)|
|[Ddb](worksheetfunction-ddb-method-excel.md)|
|[Dec2Bin](worksheetfunction-dec2bin-method-excel.md)|
|[Dec2Hex](worksheetfunction-dec2hex-method-excel.md)|
|[Dec2Oct](worksheetfunction-dec2oct-method-excel.md)|
|[Decimal](worksheetfunction-decimal-method-excel.md)|
|[Degrees](worksheetfunction-degrees-method-excel.md)|
|[Delta](worksheetfunction-delta-method-excel.md)|
|[DevSq](worksheetfunction-devsq-method-excel.md)|
|[DGet](worksheetfunction-dget-method-excel.md)|
|[Disc](worksheetfunction-disc-method-excel.md)|
|[DMax](worksheetfunction-dmax-method-excel.md)|
|[DMin](worksheetfunction-dmin-method-excel.md)|
|[Dollar](worksheetfunction-dollar-method-excel.md)|
|[DollarDe](worksheetfunction-dollarde-method-excel.md)|
|[DollarFr](worksheetfunction-dollarfr-method-excel.md)|
|[DProduct](worksheetfunction-dproduct-method-excel.md)|
|[DStDev](worksheetfunction-dstdev-method-excel.md)|
|[DStDevP](worksheetfunction-dstdevp-method-excel.md)|
|[DSum](worksheetfunction-dsum-method-excel.md)|
|[Duration](worksheetfunction-duration-method-excel.md)|
|[DVar](worksheetfunction-dvar-method-excel.md)|
|[DVarP](worksheetfunction-dvarp-method-excel.md)|
|[EDate](worksheetfunction-edate-method-excel.md)|
|[Effect](worksheetfunction-effect-method-excel.md)|
|[EncodeURL](worksheetfunction-encodeurl-method-excel.md)|
|[EoMonth](worksheetfunction-eomonth-method-excel.md)|
|[Erf](worksheetfunction-erf-method-excel.md)|
|[Erf_Precise](worksheetfunction-erf_precise-method-excel.md)|
|[ErfC](worksheetfunction-erfc-method-excel.md)|
|[ErfC_Precise](worksheetfunction-erfc_precise-method-excel.md)|
|[Even](worksheetfunction-even-method-excel.md)|
|[Expon_Dist](worksheetfunction-expon_dist-method-excel.md)|
|[ExponDist](worksheetfunction-expondist-method-excel.md)|
|[F_Dist](worksheetfunction-f_dist-method-excel.md)|
|[F_Dist_RT](worksheetfunction-f_dist_rt-method-excel.md)|
|[F_Inv](worksheetfunction-f_inv-method-excel.md)|
|[F_Inv_RT](worksheetfunction-f_inv_rt-method-excel.md)|
|[F_Test](worksheetfunction-f_test-method-excel.md)|
|[Fact](worksheetfunction-fact-method-excel.md)|
|[FactDouble](worksheetfunction-factdouble-method-excel.md)|
|[FDist](worksheetfunction-fdist-method-excel.md)|
|[FilterXML](worksheetfunction-filterxml-method-excel.md)|
|[Find](worksheetfunction-find-method-excel.md)|
|[FindB](worksheetfunction-findb-method-excel.md)|
|[FInv](worksheetfunction-finv-method-excel.md)|
|[Fisher](worksheetfunction-fisher-method-excel.md)|
|[FisherInv](worksheetfunction-fisherinv-method-excel.md)|
|[Fixed](worksheetfunction-fixed-method-excel.md)|
|[Floor](worksheetfunction-floor-method-excel.md)|
|[Floor_Math](worksheetfunction-floor_math-method-excel.md)|
|[Floor_Precise](worksheetfunction-floor_precise-method-excel.md)|
|[Forecast](worksheetfunction-forecast-method-excel.md)|
|[Frequency](worksheetfunction-frequency-method-excel.md)|
|[FTest](worksheetfunction-ftest-method-excel.md)|
|[Fv](worksheetfunction-fv-method-excel.md)|
|[FVSchedule](worksheetfunction-fvschedule-method-excel.md)|
|[Gamma](worksheetfunction-gamma-method-excel.md)|
|[Gamma_Dist](worksheetfunction-gamma_dist-method-excel.md)|
|[Gamma_Inv](worksheetfunction-gamma_inv-method-excel.md)|
|[GammaDist](worksheetfunction-gammadist-method-excel.md)|
|[GammaInv](worksheetfunction-gammainv-method-excel.md)|
|[GammaLn](worksheetfunction-gammaln-method-excel.md)|
|[GammaLn_Precise](worksheetfunction-gammaln_precise-method-excel.md)|
|[Gauss](worksheetfunction-gauss-method-excel.md)|
|[Gcd](worksheetfunction-gcd-method-excel.md)|
|[GeoMean](worksheetfunction-geomean-method-excel.md)|
|[GeStep](worksheetfunction-gestep-method-excel.md)|
|[Growth](worksheetfunction-growth-method-excel.md)|
|[HarMean](worksheetfunction-harmean-method-excel.md)|
|[Hex2Bin](worksheetfunction-hex2bin-method-excel.md)|
|[Hex2Dec](worksheetfunction-hex2dec-method-excel.md)|
|[Hex2Oct](worksheetfunction-hex2oct-method-excel.md)|
|[HLookup](worksheetfunction-hlookup-method-excel.md)|
|[HypGeom_Dist](worksheetfunction-hypgeom_dist-method-excel.md)|
|[HypGeomDist](worksheetfunction-hypgeomdist-method-excel.md)|
|[IfError](worksheetfunction-iferror-method-excel.md)|
|[IfNa](worksheetfunction-ifna-method-excel.md)|
|[ImAbs](worksheetfunction-imabs-method-excel.md)|
|[Imaginary](worksheetfunction-imaginary-method-excel.md)|
|[ImArgument](worksheetfunction-imargument-method-excel.md)|
|[ImConjugate](worksheetfunction-imconjugate-method-excel.md)|
|[ImCos](worksheetfunction-imcos-method-excel.md)|
|[ImCosh](worksheetfunction-imcosh-method-excel.md)|
|[ImCot](worksheetfunction-imcot-method-excel.md)|
|[ImCsc](worksheetfunction-imcsc-method-excel.md)|
|[ImCsch](worksheetfunction-imcsch-method-excel.md)|
|[ImDiv](worksheetfunction-imdiv-method-excel.md)|
|[ImExp](worksheetfunction-imexp-method-excel.md)|
|[ImLn](worksheetfunction-imln-method-excel.md)|
|[ImLog10](worksheetfunction-imlog10-method-excel.md)|
|[ImLog2](worksheetfunction-imlog2-method-excel.md)|
|[ImPower](worksheetfunction-impower-method-excel.md)|
|[ImProduct](worksheetfunction-improduct-method-excel.md)|
|[ImReal](worksheetfunction-imreal-method-excel.md)|
|[ImSec](worksheetfunction-imsec-method-excel.md)|
|[ImSech](worksheetfunction-imsech-method-excel.md)|
|[ImSin](worksheetfunction-imsin-method-excel.md)|
|[ImSinh](worksheetfunction-imsinh-method-excel.md)|
|[ImSqrt](worksheetfunction-imsqrt-method-excel.md)|
|[ImSub](worksheetfunction-imsub-method-excel.md)|
|[ImSum](worksheetfunction-imsum-method-excel.md)|
|[ImTan](worksheetfunction-imtan-method-excel.md)|
|[Index](worksheetfunction-index-method-excel.md)|
|[Intercept](worksheetfunction-intercept-method-excel.md)|
|[IntRate](worksheetfunction-intrate-method-excel.md)|
|[Ipmt](worksheetfunction-ipmt-method-excel.md)|
|[Irr](worksheetfunction-irr-method-excel.md)|
|[IsErr](worksheetfunction-iserr-method-excel.md)|
|[IsError](worksheetfunction-iserror-method-excel.md)|
|[IsEven](worksheetfunction-iseven-method-excel.md)|
|[IsFormula](worksheetfunction-isformula-method-excel.md)|
|[IsLogical](worksheetfunction-islogical-method-excel.md)|
|[IsNA](worksheetfunction-isna-method-excel.md)|
|[IsNonText](worksheetfunction-isnontext-method-excel.md)|
|[IsNumber](worksheetfunction-isnumber-method-excel.md)|
|[ISO_Ceiling](worksheetfunction-iso_ceiling-method-excel.md)|
|[IsOdd](worksheetfunction-isodd-method-excel.md)|
|[IsoWeekNum](worksheetfunction-isoweeknum-method-excel.md)|
|[Ispmt](worksheetfunction-ispmt-method-excel.md)|
|[IsText](worksheetfunction-istext-method-excel.md)|
|[Kurt](worksheetfunction-kurt-method-excel.md)|
|[Large](worksheetfunction-large-method-excel.md)|
|[Lcm](worksheetfunction-lcm-method-excel.md)|
|[LinEst](worksheetfunction-linest-method-excel.md)|
|[Ln](worksheetfunction-ln-method-excel.md)|
|[Log](worksheetfunction-log-method-excel.md)|
|[Log10](worksheetfunction-log10-method-excel.md)|
|[LogEst](worksheetfunction-logest-method-excel.md)|
|[LogInv](worksheetfunction-loginv-method-excel.md)|
|[LogNorm_Dist](worksheetfunction-lognorm_dist-method-excel.md)|
|[LogNorm_Inv](worksheetfunction-lognorm_inv-method-excel.md)|
|[LogNormDist](worksheetfunction-lognormdist-method-excel.md)|
|[Lookup](worksheetfunction-lookup-method-excel.md)|
|[Match](worksheetfunction-match-method-excel.md)|
|[Max](worksheetfunction-max-method-excel.md)|
|[MDeterm](worksheetfunction-mdeterm-method-excel.md)|
|[MDuration](worksheetfunction-mduration-method-excel.md)|
|[Median](worksheetfunction-median-method-excel.md)|
|[Min](worksheetfunction-min-method-excel.md)|
|[MInverse](worksheetfunction-minverse-method-excel.md)|
|[MIrr](worksheetfunction-mirr-method-excel.md)|
|[MMult](worksheetfunction-mmult-method-excel.md)|
|[Mode](worksheetfunction-mode-method-excel.md)|
|[Mode_Mult](worksheetfunction-mode_mult-method-excel.md)|
|[Mode_Sngl](worksheetfunction-mode_sngl-method-excel.md)|
|[MRound](worksheetfunction-mround-method-excel.md)|
|[MultiNomial](worksheetfunction-multinomial-method-excel.md)|
|[Munit](worksheetfunction-munit-method-excel.md)|
|[NegBinom_Dist](worksheetfunction-negbinom_dist-method-excel.md)|
|[NegBinomDist](worksheetfunction-negbinomdist-method-excel.md)|
|[NetworkDays](worksheetfunction-networkdays-method-excel.md)|
|[NetworkDays_Intl](worksheetfunction-networkdays_intl-method-excel.md)|
|[Nominal](worksheetfunction-nominal-method-excel.md)|
|[Norm_Dist](worksheetfunction-norm_dist-method-excel.md)|
|[Norm_Inv](worksheetfunction-norm_inv-method-excel.md)|
|[Norm_S_Dist](worksheetfunction-norm_s_dist-method-excel.md)|
|[Norm_S_Inv](worksheetfunction-norm_s_inv-method-excel.md)|
|[NormDist](worksheetfunction-normdist-method-excel.md)|
|[NormInv](worksheetfunction-norminv-method-excel.md)|
|[NormSDist](worksheetfunction-normsdist-method-excel.md)|
|[NormSInv](worksheetfunction-normsinv-method-excel.md)|
|[NPer](worksheetfunction-nper-method-excel.md)|
|[Npv](worksheetfunction-npv-method-excel.md)|
|[NumberValue](worksheetfunction-numbervalue-method-excel.md)|
|[Oct2Bin](worksheetfunction-oct2bin-method-excel.md)|
|[Oct2Dec](worksheetfunction-oct2dec-method-excel.md)|
|[Oct2Hex](worksheetfunction-oct2hex-method-excel.md)|
|[Odd](worksheetfunction-odd-method-excel.md)|
|[OddFPrice](worksheetfunction-oddfprice-method-excel.md)|
|[OddFYield](worksheetfunction-oddfyield-method-excel.md)|
|[OddLPrice](worksheetfunction-oddlprice-method-excel.md)|
|[OddLYield](worksheetfunction-oddlyield-method-excel.md)|
|[Or](worksheetfunction-or-method-excel.md)|
|[PDuration](worksheetfunction-pduration-method-excel.md)|
|[Pearson](worksheetfunction-pearson-method-excel.md)|
|[Percentile](worksheetfunction-percentile-method-excel.md)|
|[Percentile_Exc](worksheetfunction-percentile_exc-method-excel.md)|
|[Percentile_Inc](worksheetfunction-percentile_inc-method-excel.md)|
|[PercentRank](worksheetfunction-percentrank-method-excel.md)|
|[PercentRank_Exc](worksheetfunction-percentrank_exc-method-excel.md)|
|[PercentRank_Inc](worksheetfunction-percentrank_inc-method-excel.md)|
|[Permut](worksheetfunction-permut-method-excel.md)|
|[Permutationa](worksheetfunction-permutationa-method-excel.md)|
|[Phi](worksheetfunction-phi-method-excel.md)|
|[Phonetic](worksheetfunction-phonetic-method-excel.md)|
|[Pi](worksheetfunction-pi-method-excel.md)|
|[Pmt](worksheetfunction-pmt-method-excel.md)|
|[Poisson](worksheetfunction-poisson-method-excel.md)|
|[Poisson_Dist](worksheetfunction-poisson_dist-method-excel.md)|
|[Power](worksheetfunction-power-method-excel.md)|
|[Ppmt](worksheetfunction-ppmt-method-excel.md)|
|[Price](worksheetfunction-price-method-excel.md)|
|[PriceDisc](worksheetfunction-pricedisc-method-excel.md)|
|[PriceMat](worksheetfunction-pricemat-method-excel.md)|
|[Prob](worksheetfunction-prob-method-excel.md)|
|[Product](worksheetfunction-product-method-excel.md)|
|[Proper](worksheetfunction-proper-method-excel.md)|
|[Pv](worksheetfunction-pv-method-excel.md)|
|[Quartile](worksheetfunction-quartile-method-excel.md)|
|[Quartile_Exc](worksheetfunction-quartile_exc-method-excel.md)|
|[Quartile_Inc](worksheetfunction-quartile_inc-method-excel.md)|
|[Quotient](worksheetfunction-quotient-method-excel.md)|
|[Radians](worksheetfunction-radians-method-excel.md)|
|[RandBetween](worksheetfunction-randbetween-method-excel.md)|
|[Rank](worksheetfunction-rank-method-excel.md)|
|[Rank_Avg](worksheetfunction-rank_avg-method-excel.md)|
|[Rank_Eq](worksheetfunction-rank_eq-method-excel.md)|
|[Rate](worksheetfunction-rate-method-excel.md)|
|[Received](worksheetfunction-received-method-excel.md)|
|[Replace](worksheetfunction-replace-method-excel.md)|
|[ReplaceB](worksheetfunction-replaceb-method-excel.md)|
|[Rept](worksheetfunction-rept-method-excel.md)|
|[Roman](worksheetfunction-roman-method-excel.md)|
|[Round](worksheetfunction-round-method-excel.md)|
|[RoundDown](worksheetfunction-rounddown-method-excel.md)|
|[RoundUp](worksheetfunction-roundup-method-excel.md)|
|[Rri](worksheetfunction-rri-method-excel.md)|
|[RSq](worksheetfunction-rsq-method-excel.md)|
|[RTD](worksheetfunction-rtd-method-excel.md)|
|[Search](worksheetfunction-search-method-excel.md)|
|[SearchB](worksheetfunction-searchb-method-excel.md)|
|[Sec](worksheetfunction-sec-method-excel.md)|
|[Sech](worksheetfunction-sech-method-excel.md)|
|[SeriesSum](worksheetfunction-seriessum-method-excel.md)|
|[Sinh](worksheetfunction-sinh-method-excel.md)|
|[Skew](worksheetfunction-skew-method-excel.md)|
|[Skew_p](worksheetfunction-skew_p-method-excel.md)|
|[Sln](worksheetfunction-sln-method-excel.md)|
|[Slope](worksheetfunction-slope-method-excel.md)|
|[Small](worksheetfunction-small-method-excel.md)|
|[SqrtPi](worksheetfunction-sqrtpi-method-excel.md)|
|[Standardize](worksheetfunction-standardize-method-excel.md)|
|[StDev](worksheetfunction-stdev-method-excel.md)|
|[StDev_P](worksheetfunction-stdev_p-method-excel.md)|
|[StDev_S](worksheetfunction-stdev_s-method-excel.md)|
|[StDevP](worksheetfunction-stdevp-method-excel.md)|
|[StEyx](worksheetfunction-steyx-method-excel.md)|
|[Substitute](worksheetfunction-substitute-method-excel.md)|
|[Subtotal](worksheetfunction-subtotal-method-excel.md)|
|[Sum](worksheetfunction-sum-method-excel.md)|
|[SumIf](worksheetfunction-sumif-method-excel.md)|
|[SumIfs](worksheetfunction-sumifs-method-excel.md)|
|[SumProduct](worksheetfunction-sumproduct-method-excel.md)|
|[SumSq](worksheetfunction-sumsq-method-excel.md)|
|[SumX2MY2](worksheetfunction-sumx2my2-method-excel.md)|
|[SumX2PY2](worksheetfunction-sumx2py2-method-excel.md)|
|[SumXMY2](worksheetfunction-sumxmy2-method-excel.md)|
|[Syd](worksheetfunction-syd-method-excel.md)|
|[T_Dist](worksheetfunction-t_dist-method-excel.md)|
|[T_Dist_2T](worksheetfunction-t_dist_2t-method-excel.md)|
|[T_Dist_RT](worksheetfunction-t_dist_rt-method-excel.md)|
|[T_Inv](worksheetfunction-t_inv-method-excel.md)|
|[T_Inv_2T](worksheetfunction-t_inv_2t-method-excel.md)|
|[T_Test](worksheetfunction-t_test-method-excel.md)|
|[Tanh](worksheetfunction-tanh-method-excel.md)|
|[TBillEq](worksheetfunction-tbilleq-method-excel.md)|
|[TBillPrice](worksheetfunction-tbillprice-method-excel.md)|
|[TBillYield](worksheetfunction-tbillyield-method-excel.md)|
|[TDist](worksheetfunction-tdist-method-excel.md)|
|[Text](worksheetfunction-text-method-excel.md)|
|[TInv](worksheetfunction-tinv-method-excel.md)|
|[Transpose](worksheetfunction-transpose-method-excel.md)|
|[Trend](worksheetfunction-trend-method-excel.md)|
|[Trim](worksheetfunction-trim-method-excel.md)|
|[TrimMean](worksheetfunction-trimmean-method-excel.md)|
|[TTest](worksheetfunction-ttest-method-excel.md)|
|[Unichar](worksheetfunction-unichar-method-excel.md)|
|[Unicode](worksheetfunction-unicode-method-excel.md)|
|[USDollar](worksheetfunction-usdollar-method-excel.md)|
|[Var](worksheetfunction-var-method-excel.md)|
|[Var_P](worksheetfunction-var_p-method-excel.md)|
|[Var_S](worksheetfunction-var_s-method-excel.md)|
|[VarP](worksheetfunction-varp-method-excel.md)|
|[Vdb](worksheetfunction-vdb-method-excel.md)|
|[VLookup](worksheetfunction-vlookup-method-excel.md)|
|[WebService](worksheetfunction-webservice-method-excel.md)|
|[Weekday](worksheetfunction-weekday-method-excel.md)|
|[WeekNum](worksheetfunction-weeknum-method-excel.md)|
|[Weibull](worksheetfunction-weibull-method-excel.md)|
|[Weibull_Dist](worksheetfunction-weibull_dist-method-excel.md)|
|[WorkDay](worksheetfunction-workday-method-excel.md)|
|[WorkDay_Intl](worksheetfunction-workday_intl-method-excel.md)|
|[Xirr](worksheetfunction-xirr-method-excel.md)|
|[Xnpv](worksheetfunction-xnpv-method-excel.md)|
|[Xor](worksheetfunction-xor-method-excel.md)|
|[YearFrac](worksheetfunction-yearfrac-method-excel.md)|
|[YieldDisc](worksheetfunction-yielddisc-method-excel.md)|
|[YieldMat](worksheetfunction-yieldmat-method-excel.md)|
|[Z_Test](worksheetfunction-z_test-method-excel.md)|
|[ZTest](worksheetfunction-ztest-method-excel.md)|
|[Forecast_ETS](worksheetfunction-forecast_ets-method-excel.md)|
|[Forecast_ETS_ConfInt](worksheetfunction-forecast_ets_confint-method-excel.md)|
|[Forecast_ETS_Seasonality](worksheetfunction-forecast_ets_seasonality-method-excel.md)|
|[Forecast_ETS_STAT](worksheetfunction-forecast_ets_stat-method-excel.md)|
|[Forecast_Linear](worksheetfunction-forecast_linear-method-excel.md)|

## Properties
<a name="AboutContributor"> </a>



|**Name**|
|:-----|
|[Application](worksheetfunction-application-property-excel.md)|
|[Creator](worksheetfunction-creator-property-excel.md)|
|[Parent](worksheetfunction-parent-property-excel.md)|

## See also
<a name="AboutContributor"> </a>


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
