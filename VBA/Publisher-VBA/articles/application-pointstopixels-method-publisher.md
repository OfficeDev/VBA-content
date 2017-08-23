---
title: "Метод Application.PointsToPixels (издатель)"
keywords: vbapb10.chm131161
f1_keywords: vbapb10.chm131161
ms.prod: publisher
api_name: Publisher.Application.PointsToPixels
ms.assetid: 9c67fcae-6c93-ddae-cbad-75356e5c5084
ms.date: 06/08/2017
ms.openlocfilehash: fdd12b228dec5b0386dbdcc9397a01483a72e19b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstopixels-method-publisher"></a>Метод Application.PointsToPixels (издатель)

Преобразует измерения из точки в пикселях (1 пиксель = 0,75 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToPixels** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в пикселях.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PixelsToPoints](application-pixelstopoints-method-publisher.md)** для преобразования измерений в пикселях в пунктах.


## <a name="example"></a>Пример

В этом примере выполняется преобразование измерения в пунктах, введенный пользователем измерений в пикселях.


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
 .PointsToPixels(Value:=Val(strInput)), _ 
 "0.00") &; " pixels" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

