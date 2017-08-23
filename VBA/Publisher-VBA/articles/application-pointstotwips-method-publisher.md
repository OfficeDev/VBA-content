---
title: "Метод Application.PointsToTwips (издатель)"
keywords: vbapb10.chm131168
f1_keywords: vbapb10.chm131168
ms.prod: publisher
api_name: Publisher.Application.PointsToTwips
ms.assetid: ba928b83-f551-049e-5868-098a9837ee7b
ms.date: 06/08/2017
ms.openlocfilehash: e99b5bb0d0811d3ee15a1caf07ae3e085bfaf72e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstotwips-method-publisher"></a>Метод Application.PointsToTwips (издатель)

Преобразует измерения из точек в твипы (20 твипов = 1 пункт). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToTwips** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в твипах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[TwipsToPoints](application-twipstopoints-method-publisher.md)** для преобразования измерений в твипах в пунктах.


## <a name="example"></a>Пример

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
 .PointsToTwips(Value:=Val(strInput)), _ 
 "0.00") &; " twips" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

