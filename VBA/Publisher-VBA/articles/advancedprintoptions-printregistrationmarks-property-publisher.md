---
title: "Свойство AdvancedPrintOptions.PrintRegistrationMarks (издатель)"
keywords: vbapb10.chm7077896
f1_keywords: vbapb10.chm7077896
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.PrintRegistrationMarks
ms.assetid: 24928459-0158-b7a9-46c0-c1a6116518d5
ms.date: 06/08/2017
ms.openlocfilehash: 7c238177d5c4155069e710ad5270d379bfd372ab
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsprintregistrationmarks-property-publisher"></a>Свойство AdvancedPrintOptions.PrintRegistrationMarks (издатель)

 **Значение true** для печати регистрации помечает для указанной публикации. По умолчанию используется **значение True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintRegistrationMarks**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если выбрать любой режим, отличный от цветоделение для указанной публикации возвращает «Отказано в разрешении».

Это свойство соответствует управления **совмещения** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Совмещения используются для выравнивания (реестр) печати двух или более press формы на одной странице.

Эти метки печати за пределами публикации и можно распечатать только если размер бумаги, печать для превышает размер страницы публикации.


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

