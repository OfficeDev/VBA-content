---
title: "Метод Application.PointsToMillimeters (издатель)"
keywords: vbapb10.chm131159
f1_keywords: vbapb10.chm131159
ms.prod: publisher
api_name: Publisher.Application.PointsToMillimeters
ms.assetid: eaa9154d-1a9b-81e7-58bc-3f7bf873ab97
ms.date: 06/08/2017
ms.openlocfilehash: 688d206c23db9bb6a4c29eb09dbebd8f710d7a5e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstomillimeters-method-publisher"></a>Метод Application.PointsToMillimeters (издатель)

Преобразует измерения из точки мм (1 мм = 2.835 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToMillimeters** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в мм.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[MillimetersToPoints](application-millimeterstopoints-method-publisher.md)** для преобразования измерений в мм в пунктах.


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
 .PointsToMillimeters(Value:=Val(strInput)), _ 
 "0.00") &; " mm" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

