---
title: "Метод Application.TwipsToPoints (издатель)"
keywords: vbapb10.chm131154
f1_keywords: vbapb10.chm131154
ms.prod: publisher
api_name: Publisher.Application.TwipsToPoints
ms.assetid: 18e1c4da-1295-31a2-d66b-ab0df807b7a6
ms.date: 06/08/2017
ms.openlocfilehash: 243d965a4f7758d2cdd3264d7d77248339befc9c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationtwipstopoints-method-publisher"></a>Метод Application.TwipsToPoints (издатель)

Преобразует измерения из твипов в точки (20 твипов = 1 пункт). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TwipsToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение твип для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToTwips](application-pointstotwips-method-publisher.md)** для преобразования твипов измерения в точках.


## <a name="example"></a>Пример

В этом примере выполняется преобразование измерения в твипах, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in twips (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " twips = " _ 
 &; Format(Application _ 
 .TwipsToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

