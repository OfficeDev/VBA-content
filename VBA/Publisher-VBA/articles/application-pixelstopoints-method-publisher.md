---
title: "Метод Application.PixelsToPoints (издатель)"
keywords: vbapb10.chm131153
f1_keywords: vbapb10.chm131153
ms.prod: publisher
api_name: Publisher.Application.PixelsToPoints
ms.assetid: 5d7e453f-e962-e557-48e4-44766d0c64d9
ms.date: 06/08/2017
ms.openlocfilehash: 9fbc60a25a0b9812f05790f3c64bd3af3e6ca33e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpixelstopoints-method-publisher"></a>Метод Application.PixelsToPoints (издатель)

Преобразует измерения из точек в точках (1 пиксель = 0,75 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PixelsToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение в пикселях для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToPixels](application-pointstopixels-method-publisher.md)** для преобразования измерений в точках в пикселях.


## <a name="example"></a>Пример

В этом примере преобразует измерения в пикселах, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in pixels (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " pixels = " _ 
 &; Format(Application _ 
 .PixelsToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

