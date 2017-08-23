---
title: "Свойство AdvancedPrintOptions.PrintDensityBars (издатель)"
keywords: vbapb10.chm7077904
f1_keywords: vbapb10.chm7077904
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.PrintDensityBars
ms.assetid: b98baed0-e2ba-bf69-78e2-d60125d7f57a
ms.date: 06/08/2017
ms.openlocfilehash: b6f989d1344f66d1a3ae8e75c1d6fd3bbb420d8a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsprintdensitybars-property-publisher"></a>Свойство AdvancedPrintOptions.PrintDensityBars (издатель)

 **Значение true** для печати шкалы плотности для указанной публикации. По умолчанию используется **значение True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintDensityBars**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если выбрать любой режим, отличный от цветоделение для указанной публикации возвращает «Отказано в разрешении».

На панели плотность при выводе на печать выражаемым числом экране 10 процентов для заполнения 100 процентов. Профессиональной печати можно использовать эту панель для определения нужное время выдержки при записи формы, а также тестирования растровых в печатных страниц.

Это свойство соответствует управления **плотности** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Эти метки печати за пределами публикации и может быть при выводе на печать только в том случае, если размер бумаги, печать на превышает размер страницы публикации.


## <a name="example"></a>Пример

В следующем примере задается меток обрезки и сведения о задании для печати с публикацией. При печати публикации цветоделение дополнительные типы типографские метки также необходимо задать для печати. В этом примере предполагается, что размер бумаги, печать для больше, чем размер страницы публикации.


```vb
Sub SetPrintersMarksToPrint() 
 With ActiveDocument.AdvancedPrintOptions 
 .PrintCropMarks = True 
 .PrintJobInformation = True 
 If PrintMode = pbPrintModeSeparations Then 
 .PrintRegistrationMarks = True 
 .PrintDensityBars = True 
 .PrintColorBars = True 
 End If 
 End With 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

