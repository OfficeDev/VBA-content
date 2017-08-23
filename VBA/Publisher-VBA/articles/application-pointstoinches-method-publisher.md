---
title: "Метод Application.PointsToInches (издатель)"
keywords: vbapb10.chm131157
f1_keywords: vbapb10.chm131157
ms.prod: publisher
api_name: Publisher.Application.PointsToInches
ms.assetid: 58bfd9ce-dee7-0a14-8ec1-7e16a5e967d8
ms.date: 06/08/2017
ms.openlocfilehash: bb2d2c82657ce66c00e07c9a31cfea01b70df938
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstoinches-method-publisher"></a>Метод Application.PointsToInches (издатель)

Преобразует измерения из точки дюйма (1 = 72 точки). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToInches** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в дюймах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[InchesToPoints не была назначена](application-inchestopoints-method-publisher.md)** для преобразования измерений в дюймах в пунктах.


## <a name="example"></a>Пример

В этом примере выполняется преобразование измерения в пунктах, введенный пользователем измерений в дюймах.


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
 .PointsToInches(Value:=Val(strInput)), _ 
 "0.00") &; " in" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

