---
title: "Свойство AdvancedPrintOptions.PrintCropMarks (издатель)"
keywords: vbapb10.chm7077895
f1_keywords: vbapb10.chm7077895
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.PrintCropMarks
ms.assetid: 0b777632-572c-7080-8f4d-b97a284d04e2
ms.date: 06/08/2017
ms.openlocfilehash: c7ec0d766a3109840fbffa2ca2ed5784c4718187
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsprintcropmarks-property-publisher"></a>Свойство AdvancedPrintOptions.PrintCropMarks (издатель)

 **Значение true** для печати обрезки помечает для указанной публикации. По умолчанию используется **значение True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintCropMarks**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство соответствует управления **меток обрезки** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Обрезка метки используются в качестве руководства при публикации для печати обрезается до требуемого размера.

Эти метки печати за пределами публикации и можно распечатать только если размер листа, печать для превышает размер страницы публикации.


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

