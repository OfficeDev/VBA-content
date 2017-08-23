---
title: "Метод Application.PointsToEmus (издатель)"
keywords: vbapb10.chm131156
f1_keywords: vbapb10.chm131156
ms.prod: publisher
api_name: Publisher.Application.PointsToEmus
ms.assetid: cb3f0bb9-fa0d-d967-9294-081a369c2c4e
ms.date: 06/08/2017
ms.openlocfilehash: ddac987710af1aa5719c8276c20bbd8f6796c9a7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstoemus-method-publisher"></a>Метод Application.PointsToEmus (издатель)

Преобразует измерения из точки emus (12700 emus = 1 пункт). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToEmus** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в emus.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[EmusToPoints](application-emustopoints-method-publisher.md)** для преобразования измерений в emus в пунктах.


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
 .PointsToEmus(Value:=Val(strInput)), _ 
 "0.00") &; " emus" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

