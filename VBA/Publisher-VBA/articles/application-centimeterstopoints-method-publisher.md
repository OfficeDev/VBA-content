---
title: "Метод Application.CentimetersToPoints (издатель)"
keywords: vbapb10.chm131141
f1_keywords: vbapb10.chm131141
ms.prod: publisher
api_name: Publisher.Application.CentimetersToPoints
ms.assetid: 6eda6692-ea9a-c4ad-6991-066fdc23bd2c
ms.date: 06/08/2017
ms.openlocfilehash: abe29dd5d41921a672676e34358977f41cb4cbb0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationcentimeterstopoints-method-publisher"></a>Метод Application.CentimetersToPoints (издатель)

Преобразует измерения из см в точках (1 cm = 28.35 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CentimetersToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение сантиметр для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToCentimeters](application-pointstocentimeters-method-publisher.md)** для преобразования значения в точках см.


## <a name="example"></a>Пример

В этом примере выполняется преобразование размеры в см, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in centimeters (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " cm = " _ 
 &; Format(Application _ 
 .CentimetersToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

