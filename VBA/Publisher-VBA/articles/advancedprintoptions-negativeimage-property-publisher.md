---
title: "Свойство AdvancedPrintOptions.NegativeImage (издатель)"
keywords: vbapb10.chm7077893
f1_keywords: vbapb10.chm7077893
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.NegativeImage
ms.assetid: 32a524ce-da31-8dfa-3286-c5d9c74367ca
ms.date: 06/08/2017
ms.openlocfilehash: cec9cd70f47fcfa96e1a29e457cdfd03b8e44366
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsnegativeimage-property-publisher"></a>Свойство AdvancedPrintOptions.NegativeImage (издатель)

 **Значение true** для печати негативного изображения указанной публикации. Значение по умолчанию — **False**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NegativeImage**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство доступно только в том случае, если active принтер PostScript принтера. Возвращает ошибку времени выполнения, если указан другие принтера. Используйте свойство **[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** для определения, если указанный принтер PostScript принтера.

Данное свойство применяется для будущего экземпляры Microsoft Publisher и сохраняются как параметр приложения.

Это свойство соответствует **отрицательные** рисунка на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .


## <a name="example"></a>Пример

Следующий пример определяет, является ли активного принтера PostScript принтера. Если он установлен, active публикация предназначена для печати, как зеркальное копирование по горизонтали и по вертикали, отрицательные изображение самого себя.


```vb
Sub PrepToPrintToFilmOnImagesetter() 
 
With ActiveDocument.AdvancedPrintOptions 
 If .IsPostscriptPrinter = True Then 
 .HorizontalFlip = True 
 .VerticalFlip = True 
 .NegativeImage = True 
 End If 
End With 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

