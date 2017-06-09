---
title: PjDateLabel Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjDateLabel
ms.assetid: ece69c4d-35fc-a795-8acb-1ff79df9fe1c
ms.date: 06/08/2017
---


# PjDateLabel Enumeration (Project)

Contains constants that specify the display format for date and time labels in a timescale.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjDay_ddd**|19|Examples: Mon, Tue. Requires the time unit to be  **pjTimescaleDays**.|
|**pjDay_ddd_dd**|105|Examples: Mon 30, Tue 1|
|**pjDay_ddd_m_dd**|112|Examples: Mon S 30, Tue O 1|
|**pjDay_ddd_mm_dd**|108|Examples: Mon 9/30, Tue 10/1|
|**pjDay_ddd_mm_dd_yy**|52|Examples: Mon 9/30/02, Tue 10/1/02|
|**pjDay_ddd_mmm_dd**|23|Examples: Mon Sep 30, Tue Oct 1|
|**pjDay_ddd_mmm_dd_yyy**|22|Examples: Mon September 30 '02, Tue October 1 '02|
|**pjDay_ddd_mmmm_dd**|111|Examples: Mon September 30, Tue October 1|
|**pjDay_dddd**|18|Examples: Tuesday, Wednesday.|
|**pjDay_ddi**|119|Examples: Mo, Tu|
|**pjDay_ddi_dd**|106|Examples: Mo 30, Tu 1|
|**pjDay_ddi_m_dd**|113|Examples: Mo S 30, Tu O 1|
|**pjDay_ddi_mm_dd**|109|Examples: Mo 9/30, Tu 10/1|
|**pjDay_di**|20|Examples: M, T|
|**pjDay_di_dd**|107|Examples: M 30, T 1|
|**pjDay_di_m_dd**|114|Examples: M S 30, T O 1|
|**pjDay_di_mm_dd**|110|Examples: M 9/30, T 10/1|
|**pjDay_didd**|121|Examples: M30, T1|
|**pjDay_m_dd**|115|Examples: S 30, O 1|
|**pjDay_mm_dd**|27|Examples: 9/30, 10/1|
|**pjDay_mm_dd_yy**|26|Examples: 9/30/02, 10/1/02|
|**pjDay_mmm_dd**|25|Examples: Sep 30, Oct 1|
|**pjDay_mmm_dd_yyy**|24|Examples: Sep 30 '02, Oct 10 '02|
|**pjDayFromEnd_Day_dd**|41|Examples: Day 2, Day 1, Day -1, Day -2 from the project end.|
|**pjDayFromEnd_dd**|54|Examples: 2, 1, -1, -2|
|**pjDayFromEnd_Ddd**|53|Examples: D2, D1, D-1, D-2|
|**pjDayFromStart_Day_dd**|40|Examples: Day -2, Day -1, Day 1, Day 2 from the project start.|
|**pjDayFromStart_dd**|56|Examples: -2, -1, 1, 2|
|**pjDayFromStart_Ddd**|55|Examples: D-2, D-1, D1, D2|
|**pjDayOfMonth_dd**|21|Examples: 30, 1|
|**pjDayOfYear_dd**|118|Examples: 77, 78|
|**pjDayOfYear_dd_yyy**|116|Examples: 77 '10, 78 '10|
|**pjDayOfYear_dd_yyyy**|117|Examples: 77 2010, 78 2010|
|**pjHalfYear_h**|128|Examples: 1, 2. Requires the time unit to be  **pjTimescaleHalfYears**.|
|**pjHalfYear_Hh**|127|Examples: H1, H2|
|**pjHalfYear_Hh_yyy**|126|Examples: H1 '10, H2 '10|
|**pjHalfYear_hhh_Half**|123|Examples: 1st Half, 2nd Half|
|**pjHalfYear_hHyy**|129|Examples: 1H10, 2H10|
|**pjHalfYear_Hlf_h**|125|Examples: Half 1, Half 2|
|**pjHalfYear_Hlf_h_yyyy**|124|Examples: Half 1, 2010; Half 2, 2010|
|**pjHalfYearFromEnd_h**|135|Examples: 2, 1, -1, -2. Half years from the project end date.|
|**pjHalfYearFromEnd_Half_h**|133|Examples: Half 2, Half 1, Half -1, Half -2 |
|**pjHalfYearFromEnd_Hh**|134|Examples: H2, H1, H-1, H-2|
|**pjHalfYearFromStart_h**|132|Examples: -2, -1, 1, 2. Half years from the project start date.|
|**pjHalfYearFromStart_Half_h**|130|Examples: Half -2, Half -1, Half 1, Half 2|
|**pjHalfYearFromStart_Hh**|131|Examples: H-2, H-1, H1, H2|
|**pjHour_ddd_mmm_dd_hhAM**|28|Examples: Wed Mar 18, 8 AM; Wed Mar 18, 9 AM. Requires the time unit to be  **pjTimescaleHours**.|
|**pjHour_hh**|32|Examples: 8, 9, 10, 11|
|**pjHour_hh_mmAM**|30|Examples: 8:00 AM, 9:00 AM|
|**pjHour_hhAM**|31|Examples: 8AM, 9AM|
|**pjHour_mm_dd_hhAM**|120|Examples: 3/18 8 AM, 3/18 9 AM|
|**pjHour_mmm_dd_hhAM**|29|Examples: Mar 18, 8 AM; Mar 18, 9 AM|
|**pjHourFromEnd_hh**|77|Examples: 3, 2, 1, -1, -2 hours from the project end.|
|**pjHourFromEnd_Hhh**|76|Examples: H3, H2, H1, H-1, H-2|
|**pjHourFromEnd_Hour_hh**|39|Examples: Hour 3, Hour 2, Hour 1, Hour -1, Hour -2|
|**pjHourFromStart_hh**|79|Examples: -2, -1, 1, 2, 3 hours from the project start.|
|**pjHourFromStart_Hhh**|78|Examples: H-2, H-1, H1, H2, H3|
|**pjHourFromStart_Hour_hh**|38|Examples: Hour -2, Hour -1, Hour 1, Hour 2, Hour 3|
|**pjMinute_hh_mmAM**|33|Examples: 8:00 AM, 8:01 AM, 8:02 AM. Requires the time unit to be  **pjTimescaleMinutes**.|
|**pjMinute_mm**|34|Examples: 0, 1, 2, ..., 59 minutes|
|**pjMinuteFromEnd_Minute_mm**|37|Examples: Minute 181, Minute 180, ..., Minute 1, Minute -1 from the project end.|
|**pjMinuteFromEnd_mm**|81|Examples: 181, 180, ..., 1, -1|
|**pjMinuteFromEnd_Mmm**|80|Examples: M181, M180, ..., M1, M-1|
|**pjMinuteFromStart_Minute_mm**|36|Examples: Minute -2, Minute -1, Minute 1, ... Minute 180 from the project start.|
|**pjMinuteFromStart_mm**|83|Examples: -2, -1, 1, ..., 180|
|**pjMinuteFromStart_Mmm**|82|Examples: M-2, M-1, M1, ..., M180|
|**pjMonth_m**|11|Examples: M, A, M, J, J. Requires the time unit to be  **pjTimescaleMonths**.|
|**pjMonth_mm**|57|Examples: 11, 12, 1, 2|
|**pjMonth_mm_yy**|86|Examples: 3/10, 4/10, 5/10|
|**pjMonth_mm_yyy**|85|Examples: 3 '10, 4 '10, 5 '10|
|**pjMonth_mmm**|10|Examples: Mar, Apr, May|
|**pjMonth_mmm_yyy**|8|Examples: Mar '10, Apr '10, May '10|
|**pjMonth_mmmm**|9|Examples: March, April, May|
|**pjMonth_mmmm_yyyy**|7|Examples: March 2010, April 2010, May 2010|
|**pjMonthFromEnd_mm**|59|Examples: 2, 1, -1, -2 months from the project end.|
|**pjMonthFromEnd_Mmm**|58|Examples: M2, M1, M-1, M-2|
|**pjMonthFromEnd_Month_mm**|45|Examples: Month 2, Month 1, Month -1, Month -2|
|**pjMonthFromStart_mm**|61|Examples: -2, -2, 1, 2 months from the project start.|
|**pjMonthFromStart_Mmm**|60|Examples: M-2, M-1, M1, M2|
|**pjMonthFromStart_Month_mm**|44|Examples: Month -2, Month -1, Month 1, Month 2|
|**pjQuarter_q**|62|Examples: 3, 4, 1. Requires the time unit to be  **pjTimescaleQuarters**.|
|**pjQuarter_Qq**|6|Examples: Q3, Q4, Q1|
|**pjQuarter_Qq_yyy**|4|Examples: Q3 '10, Q4 '10, Q1 '11|
|**pjQuarter_qqq_Quarter**|2|Examples: 3rd Quarter, 4th Quarter, 1st Quarter|
|**pjQuarter_qQyy**|51|Examples: 3Q10, 4Q10, 1Q11|
|**pjQuarter_Qtr_q**|5|Examples: Qtr 3, Qtr 4, Qtr1|
|**pjQuarter_Qtr_q_yyyy**|3|Examples: Qtr3, 2010; Qtr4, 2010; Qtr1, 2011|
|**pjQuarterFromEnd_q**|64|Examples: 5, 4, 3, 2, 1, -1 quarters from the project end.|
|**pjQuarterFromEnd_Qq**|63|Examples: Q5, Q4, Q3, Q2, Q1, Q-1|
|**pjQuarterFromEnd_Quarter_q**|47|Examples: Quarter 5, Quarter 4, Quarter 3, Quarter 2, Quarter 1, Quarter -1|
|**pjQuarterFromStart_q**|66|Examples: -5, -4, -3, -2, -1, 1 quarters from the project start.|
|**pjQuarterFromStart_Qq**|65|Examples: Q-5, Q-4, Q-3, Q-2, Q-1, Q1|
|**pjQuarterFromStart_Quarter_q**|46|Examples: Quarter -5, Quarter -4, Quarter -3, Quarter -2, Quarter -1, Quarter 1|
|**pjThirdsOfMonths_dd**|136|Examples: 1, 11, 21, 1. Requires the time unit to be  **pjTimescaleThirdsOfMonths**.|
|**pjThirdsOfMonths_ddd**|137|Examples: B, M, E, B|
|**pjThirdsOfMonths_dddd**|138|Examples: Beginning, Middle, End, Beginning|
|**pjThirdsOfMonths_mm_dd**|139|Examples: 3/1, 3/11, 3/21, 4/1|
|**pjThirdsOfMonths_mm_dd_yy**|145|Examples: 3/1/10, 3/11/10, 3/21/10, 4/1/10|
|**pjThirdsOfMonths_mm_ddd**|140|Examples: 3/B, 3/M, 3/E, 4/B|
|**pjThirdsOfMonths_mm_ddd_yy**|146|Examples: 3/B/10, 3/M/10, 3/E/10, 4/B/10|
|**pjThirdsOfMonths_mmm_dd**|142|Examples: Mar 1, Mar 11, Mar 21, Apr 1|
|**pjThirdsOfMonths_mmm_dd_yy**|147|Examples: Mar 1, '10; Mar 11, '10; Mar 21, '10; Apr 1, 10|
|**pjThirdsOfMonths_mmm_ddd**|143|Examples: Mar B, Mar M, Mar E, Apr B|
|**pjThirdsOfMonths_mmm_ddd_yy**|148|Examples: Mar B, '10; Mar M, '10; Mar E, '10; Apr B '10|
|**pjThirdsOfMonths_mmmm_dd**|144|Examples: March 1, March 11, March 21, April 1|
|**pjThirdsOfMonths_mmmm_dd_yyyy**|149|Examples: March 1, 2010; March 11, 2010; March 21, 2010; April 1, 2010|
|**pjThirdsOfMonths_mmmm_dddd**|141|Examples: March Beginning, March Middle, March End, April Beginning|
|**pjThirdsOfMonths_mmmm_dddd_yyyy**|150|Examples: March Beginning, 2010; March Middle, 2010; March End, 2010; April Beginning, 2010|
|**pjWeek_ddd_dd**|88|Examples: Sun 21, Sun 28, Sun 4. Requires the time unit to be  **pjTimescaleWeeks**.|
|**pjWeek_ddd_m_dd**|97|Examples: Sun M 21, Sun M 28, Sun A 4|
|**pjWeek_ddd_mm_dd**|90|Examples: Sun 3/21, Sun 3/28, Sun 4/4|
|**pjWeek_ddd_mm_dd_yy**|100|Examples: Sun 3/21/10, Sun 3/28/10, Sun 4/4/10|
|**pjWeek_ddd_mmm_dd**|93|Examples: Sun Mar 21, Sun Mar 28, Sun Apr 4|
|**pjWeek_ddd_mmm_dd_yyy**|101|Examples: Sun Mar 21, '10; Sun Mar 28, '10; Sun Apr 4, '10|
|**pjWeek_ddd_mmmm_dd**|96|Examples: Sun Mar 21, Sun March 28, Sun Apr 4|
|**pjWeek_ddd_mmmm_dd_yyy**|102|Examples: Sun Mar 21, '10; Sun March 28, '10; Sun Apr 4, '10|
|**pjWeek_ddd_ww**|103|Examples: Sun 12, Sun 13, Sun 14|
|**pjWeek_ddi_m_dd**|98|Examples: Sun M 21, Sun M 28, Sun A 4|
|**pjWeek_ddi_mm_dd**|91|Examples: Su 3/21. Su 3/28, Su 4/4|
|**pjWeek_ddi_mmm_dd**|94|Examples: Su Mar 21, Su Mar 28, Su Apr 4|
|**pjWeek_di_m_dd**|99|Examples: S M 21, S M 28, S A 4|
|**pjWeek_di_mm_dd**|92|Examples: S 3/21, S 3/28, S 4/4|
|**pjWeek_di_mmm_dd**|95|Examples: S Mar 21, S Mar 28, S Apr 4|
|**pjWeek_m_dd**|89|Examples: M21, M28, A 4|
|**pjWeek_mm_dd**|17|Examples: 3/21, 3/28, 4/4|
|**pjWeek_mm_dd_yy**|16|Examples: 3/21/10. 3/28/10, 4/4/10|
|**pjWeek_mmm_dd**|15|Examples: Mar 21, Mar 28, Apr 4|
|**pjWeek_mmm_dd_yyy**|13|Examples: Mar 21, '10; Mar 28, '10; Apr 4, '10|
|**pjWeek_mmmm_dd**|14|Examples: March 21, March 28, April 4|
|**pjWeek_mmmm_dd_yyyy**|12|Examples: March 21, 2010; March 28, 2010; April 4, 2010|
|**pjWeekDayOfMonth_dd**|87|Examples: 21, 28, 4|
|**pjWeekFromEnd_Week_ww**|43|Examples: Week 2, Week 1, Week -1 from the project end.|
|**pjWeekFromEnd_ww**|68|Examples: 2, 1, -1|
|**pjWeekFromEnd_Www**|67|Examples: W2, W1, W-1|
|**pjWeekFromStart_Week_ww**|42|Examples: Week -1, Week 1, Week 2 from the project start.|
|**pjWeekFromStart_ww**|70|Examples: -1, 1, 2|
|**pjWeekFromStart_Www**|69|Examples: W-1, W1, W2|
|**pjWeekNumber_dd_ww**|104|Examples: 1 12, 1 13, 1 14 (day 1 of week 12, day 1 of week 13, and so forth)|
|**pjWeekNumber_ww**|50|Examples: 12, 13, 14|
|**pjYear_yy**|75|Examples: 10, 11, 12. Requires the time unit to be  **pjTimescaleYears**.|
|**pjYear_yyy**|1|Examples: '10, '11, '12|
|**pjYear_yyyy**|0|Examples: 2010, 2011, 2012|
|**pjYearFromEnd_Year_yy**|49|Examples: Year 2, Year 1, Year -1 from the project end.|
|**pjYearFromEnd_yy**|72|Examples: 2, 1, -1|
|**pjYearFromEnd_Yyy**|71|Examples: Y2, Y1, Y-1|
|**pjYearFromStart_Year_yy**|48|Examples: Year -1, Year 1, Year 2 from the project start.|
|**pjYearFromStart_yy**|74|Examples: -1, 1, 2|
|**pjYearFromStart_Yyy**|73|Examples: Y-1, Y1, Y2|

