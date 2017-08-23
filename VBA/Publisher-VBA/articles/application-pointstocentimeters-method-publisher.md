---
title: "Метод Application.PointsToCentimeters (издатель)"
keywords: vbapb10.chm131155
f1_keywords: vbapb10.chm131155
ms.prod: publisher
api_name: Publisher.Application.PointsToCentimeters
ms.assetid: 9a734d3d-78d2-1e27-63b3-2ad1074e16c1
ms.date: 06/08/2017
ms.openlocfilehash: f1c4bd11408bb887a235bec200bbf2a653190032
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# Метод Application.PointsToCentimeters (издатель)

Преобразует измерения из точки см (1 cm = 28.35 точек). Возвращает преобразованные измерения как **один**.


## Синтаксис

 _выражение_. **PointsToCentimeters** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в см.|

### Возвращаемое значение

Один


## Заметки

Используйте метод **[CentimetersToPoints](application-centimeterstopoints-method-publisher.md)** для преобразования измерений в см в пунктах.


## Пример

В этом примере выполняется преобразование измерения в пунктах, введенный пользователем измерений в см.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in points (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " points = " _ 
 &; Format(Application _ 
 .PointsToCentimeters(Value:=Val(strInput)), _ 
 "0.00") &; " cm" 
 
 MsgBox strOutput 
Loop
```


## См. также


#### Основные понятия


 [Объект приложения](application-object-publisher.md)

