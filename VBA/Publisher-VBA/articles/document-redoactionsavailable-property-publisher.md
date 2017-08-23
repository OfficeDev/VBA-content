---
title: "Свойство Document.RedoActionsAvailable (издатель)"
keywords: vbapb10.chm196727
f1_keywords: vbapb10.chm196727
ms.prod: publisher
api_name: Publisher.Document.RedoActionsAvailable
ms.assetid: 9af11772-e807-730a-89a0-da06e979f834
ms.date: 06/08/2017
ms.openlocfilehash: d12ad5c38acddc7b81c2851075b4b2a02f81f5ca
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentredoactionsavailable-property-publisher"></a>Свойство Document.RedoActionsAvailable (издатель)

Возвращает число действий, доступных в стеке повтора. Только для чтения **времени**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RedoActionsAvailable**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В следующем примере добавляется прямоугольника, который содержит фрагмент текста на страницу четвертый active публикации. Некоторые свойства шрифта и текст надписи зависят от выбора. Затем запускается для определения, является ли шрифт в элементе frame текст Courier. Если это так, метод **[Отменить](document-undo-method-publisher.md)** используется со значением свойства **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** , передается как параметр для указания, что все предыдущие действия быть отменены.

Метод **[Повторить](document-redo-method-publisher.md)** нажмите используется со значением свойства **RedoActionsAvailable** минус 2, передается как параметр для возврата всех действий, за исключением последние два. Новый шрифт указан текст в текстовой рамки, в дополнение к новый текст.

В этом примере предполагается, что активный документ содержит по крайней мере четыре страницы.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 12 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 
 If thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Courier" Then 
 ' The Undo method specifies that all undoable actions be undone. 
 .Undo (.UndoActionsAvailable) 
 ' The Redo method uses RedoActionsAvailable - 2 to specify that 
 ' all redoable actions be redone except for the last two actions. 
 ' The last two actions that are not redone are setting 
 ' .Font.Name and .Text. 
 .Redo (.RedoActionsAvailable - 2) 
 With theShape.TextFrame.TextRange 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End If 
End With
```


