---
title: WorksheetFunction.Binom_Dist Method (Excel)
keywords: vbaxl10.chm137414
f1_keywords:
- vbaxl10.chm137414
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Binom_Dist
ms.assetid: acd56b17-5304-0095-2696-11797df056ca
ms.date: 06/08/2017
---


# WorksheetFunction.Binom_Dist Method (Excel)

Returns the individual term binomial distribution probability.


## Syntax

 _expression_ . **Binom_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number_s - the number of successes in trials.|
| _Arg2_|Required| **Double**|Trials - the number of independent trials.|
| _Arg3_|Required| **Double**|Probability_s - the probability of success on each trial.|
| _Arg4_|Required| **Boolean**|Cumulative - a logical value that determines the form of the function. If cumulative is  **True** , then the **Binom_Dist** method returns the cumulative distribution function, which is the probability that there are at most number_s successes; if **False** , it returns the probability mass function, which is the probability that there are number_s successes.|

### Return Value

 **Double**


## Remarks

Use the  **Binom_Dist** method in problems with a fixed number of tests or trials, when the outcomes of any trial are only success or failure, when trials are independent, and when the probability of success is constant throughout the experiment. For example, the **Binom_Dist** method can calculate the probability that two of the next three babies born are male.


- Number_s and trials are truncated to integers.
    
- If number_s, trials, or probability_s is nonnumeric, the  **Binom_Dist** method generates an error.
    
- If number_s < 0 or number_s > trials, the  **Binom_Dist** method generates an error.
    
- If probability_s < 0 or probability_s > 1, the  **Binom_Dist** method generates an error.
    
- The binomial probability mass function is:
![BPM function](images/awfbnmd1_ZA06051113.gif)where: 
![BPM function](images/awfbnmd2_ZA06051114.gif)is COMBIN(n,x). The cumulative binomial distribution is: 
![BPM function](images/awfbnmd3_ZA06051115.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

