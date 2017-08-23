---
title: "Метод BorderArtFormat.RevertToDefaultWeight (издатель)"
keywords: vbapb10.chm7602180
f1_keywords: vbapb10.chm7602180
ms.prod: publisher
api_name: Publisher.BorderArtFormat.RevertToDefaultWeight
ms.assetid: 3e46637f-3fce-3346-9193-063be40844bd
ms.date: 06/08/2017
ms.openlocfilehash: 6f802970f5985864e173181448972ae3cd18f6d4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatreverttodefaultweight-method-publisher"></a>Метод BorderArtFormat.RevertToDefaultWeight (издатель)

Задает Узорные указанные форму обратно в его толщина по умолчанию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RevertToDefaultWeight**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


## <a name="remarks"></a>Заметки

Метод **RevertToDefaultWeight** имеет тот же эффект, как элемент управления **всегда применяются на размер по умолчанию** в диалоговом окне **Узорные** .

Используйте свойство **[Вес](borderartformat-weight-property-publisher.md)** объекта **[BorderArtFormat](borderartformat-object-publisher.md)** Установка указанного Узорные толщины используемый по умолчанию.


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

