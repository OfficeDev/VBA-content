---
title: "Метод Application.InchesToPoints (издатель)"
keywords: vbapb10.chm131143
f1_keywords: vbapb10.chm131143
ms.prod: publisher
api_name: Publisher.Application.InchesToPoints
ms.assetid: 32c8740f-ad14-c947-b960-500378a5873d
ms.date: 06/08/2017
ms.openlocfilehash: 621dd3202f14a5ad59072d82b53678a1fe562c93
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationinchestopoints-method-publisher"></a>Метод Application.InchesToPoints (издатель)

Преобразует измерения из дюймов точек (1 дюйм = 72 точки). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InchesToPoints не была назначена** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение дюйма для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToInches](application-pointstoinches-method-publisher.md)** для преобразования в дюймах измерения в точках.


## <a name="example"></a>Пример

В этом примере выполняется преобразование измерения в дюймах, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in inches (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " in = " _ 
 &; Format(Application _ 
 .InchesToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

