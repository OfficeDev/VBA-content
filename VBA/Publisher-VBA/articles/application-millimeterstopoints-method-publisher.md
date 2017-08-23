---
title: "Метод Application.MillimetersToPoints (издатель)"
keywords: vbapb10.chm131145
f1_keywords: vbapb10.chm131145
ms.prod: publisher
api_name: Publisher.Application.MillimetersToPoints
ms.assetid: 40ec9abd-cc1e-9f44-3312-d6689b4822e4
ms.date: 06/08/2017
ms.openlocfilehash: a3ecca08e4761e94eb4e2f14eb04c6a5b56ca62a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmillimeterstopoints-method-publisher"></a>Метод Application.MillimetersToPoints (издатель)

Преобразует измерения из мм точек (1 мм = 2.835 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MillimetersToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение мм для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToMillimeters](application-pointstomillimeters-method-publisher.md)** для преобразования значения в пунктах мм.


## <a name="example"></a>Пример

В этом примере выполняется преобразование размеры в мм, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in millimeters (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " mm = " _ 
 &; Format(Application _ 
 .Mill imetersToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

