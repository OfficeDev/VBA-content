---
title: "Свойство AdvancedPrintOptions.HorizontalFlip (издатель)"
keywords: vbapb10.chm7077892
f1_keywords: vbapb10.chm7077892
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.HorizontalFlip
ms.assetid: afb61c80-4706-8602-e78a-be35e2966c8c
ms.date: 06/08/2017
ms.openlocfilehash: 31041cf08bebef4853bc32a814631a23e3328c93
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionshorizontalflip-property-publisher"></a>Свойство AdvancedPrintOptions.HorizontalFlip (издатель)

 **Значение true** для печати по горизонтали зеркальную указанной публикации. Значение по умолчанию — **False**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalFlip**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


## <a name="remarks"></a>Заметки

Это свойство доступно только в том случае, если active печати используется принтер PostScript. Возвращает ошибку времени выполнения, если указан другие принтера. Используйте свойство **[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** для определения, если указанный принтер PostScript принтера.

Данное свойство применяется для будущего экземпляры Microsoft Publisher и сохраняются как параметр приложения.

Это свойство соответствует управления **отражение по горизонтали** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Это свойство используется в основном при печати фильма фотонаборным, чтобы правильного считывания изображения, когда на лицевой стороне фильма не работает (как при записи на форме press).


## <a name="example"></a>Пример

Следующий пример определяет, является ли активного принтера PostScript принтера. Если он установлен, активных публикация предназначена для печати как по горизонтали зеркальное копирование и по вертикали зеркальное копирование, отрицательные изображение самого себя.


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

