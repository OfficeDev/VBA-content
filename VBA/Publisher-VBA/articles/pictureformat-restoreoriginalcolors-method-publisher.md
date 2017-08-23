---
title: "Метод PictureFormat.RestoreOriginalColors (издатель)"
keywords: vbapb10.chm3604800
f1_keywords: vbapb10.chm3604800
ms.prod: publisher
api_name: Publisher.PictureFormat.RestoreOriginalColors
ms.assetid: 13a0d09f-f809-a1ca-73d9-313ea293d56a
ms.date: 06/08/2017
ms.openlocfilehash: e1d846bd990abf91f329e9f2fc14868c9ca91008
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatrestoreoriginalcolors-method-publisher"></a>Метод PictureFormat.RestoreOriginalColors (издатель)

Восстанавливает исходные цвета изображения, который был перекрашены.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RestoreOriginalColors**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


## <a name="remarks"></a>Заметки

Метод **RestoreOriginalColors** соответствует **Вернуть исходные цвета** кнопки в диалоговом окне **Перекрашивание рисунков** . (В меню **Формат** выберите пункт **изображение**и нажмите кнопку **изменить цвет**)


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **RestoreOriginalColors** для восстановления исходного цвета изображения, который был перекрашены с помощью ** [PictureFormat.Recolor](pictureformat-recolor-method-publisher.md)** метод. Он recolors первой фигуры в коллекции **фигур** на первой странице публикации и затем восстанавливает исходные цвета.

В данном примере для работы Перекрашенные фигуры значения изображения или объекта, который представляет изображение.




```vb
Public Sub RestoreOriginalColors_Example() 
 
 Dim pubPictureFormat As Publisher.PictureFormat 
 Dim pubShape As Publisher.Shape 
 Dim pubColorFormat As Publisher.ColorFormat 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 Set pubPictureFormat = pubShape.PictureFormat 
 Set pubColorFormat = pubShape.Fill.BackColor 
 
 pubPictureFormat.Recolor pubColorFormat, msoTrue 
 MsgBox "Picture was recolored." 
 pubPictureFormat.RestoreOriginalColors 
 MsgBox "Original colors in picture were restored." 
 
 
End Sub
```


