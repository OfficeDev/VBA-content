---
title: "Свойство AdvancedPrintOptions.PrintJobInformation (издатель)"
keywords: vbapb10.chm7077897
f1_keywords: vbapb10.chm7077897
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.PrintJobInformation
ms.assetid: c4494804-6dfa-8647-a72d-591f90624c1c
ms.date: 06/08/2017
ms.openlocfilehash: 855ff02c2bbe8e465e3254c944f90936bd89fe00
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsprintjobinformation-property-publisher"></a>Свойство AdvancedPrintOptions.PrintJobInformation (издатель)

 **Значение true** для печати сведений о задании печати на каждой форме. По умолчанию используется **значение True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintJobInformation**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Свойство **PrintJobInformation** можно задать независимо от того, режим, выбранные для публикации. Тем не менее, он будет проигнорирован (и печать не сведения о задании) при печати используется режим как совмещенных RGB.

Сведения о задании включает в себя имя файла распечатанных публикации, дату распечатки, номер страницы и цветовой рукописного ввода форму для (голубой, пурпурный, желтый, черный или плашечный цвет).

Это свойство соответствует элементу управления **сведения о задании** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Эти метки печати за пределами публикации и может быть при выводе на печать только в том случае, если размер бумаги, печать для превышает размер страницы публикации.


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

