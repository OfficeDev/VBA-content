---
title: "Метод BorderArtFormat.RevertToOriginalColor (издатель)"
keywords: vbapb10.chm7602192
f1_keywords: vbapb10.chm7602192
ms.prod: publisher
api_name: Publisher.BorderArtFormat.RevertToOriginalColor
ms.assetid: 6b966576-eac4-3e55-ffdc-c064341474c0
ms.date: 06/08/2017
ms.openlocfilehash: 72a761d43629f30e9b2694ac2b773caf1ed7da97
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatreverttooriginalcolor-method-publisher"></a>Метод BorderArtFormat.RevertToOriginalColor (издатель)

Задает Узорные указанные форму обратно в его цвет по умолчанию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RevertToOriginalColor**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


## <a name="remarks"></a>Заметки

Метод **RevertToOriginalColor** имеет тот же эффект, как выбор **по умолчанию** на элемент управления **цвета** **Формат < фигуры&gt; ** диалоговое окно.

Используйте свойство **[цвет](borderartformat-color-property-publisher.md)** объекта **[BorderArtFormat](borderartformat-object-publisher.md)** Установка Узорные цвет, отличный от исходного цвета.


## <a name="example"></a>Пример

Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы активных документов. Если существует Узорные его вес задано значение по умолчанию толщина и исходный цвет.


```vb
Sub RestoreBorderArtDefaults() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .RevertToDefaultWeight 
 .RevertToOriginalColor 
 End If 
 End With 
 Next anyShape 
Next anyPage 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

