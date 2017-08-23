---
title: "Свойство WebNavigationBarSet.AutoUpdate (издатель)"
keywords: vbapb10.chm8519689
f1_keywords: vbapb10.chm8519689
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.AutoUpdate
ms.assetid: b9ce8dde-c09f-6fe9-6935-cb4903a17b85
ms.date: 06/08/2017
ms.openlocfilehash: 88bd8d9d3b2ceb3bfa7f10dde3402be6362ceb80
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetautoupdate-property-publisher"></a>Свойство WebNavigationBarSet.AutoUpdate (издатель)

 **Значение true,** Если все страницы будет добавлен в указанный набор панель навигации Web и, что добавление новой страницы приведут к обновлению панели навигации с помощью соответствующего элемента. Страница должна иметь **AddHyperlinkToWebNavbar** , задайте значение **True** или свойство **WebPageOptions.IncludePageOnNewWebNavigationBars** значение **True** для добавления или обновления в рамках указанного **WebNavigationBarSet**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Автоматическое обновление**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство определяет ли существующие страницы в публикации будет добавлен в область навигации и, если добавлено страниц также будут обновлены. Эти страницы должны быть отмечены с **AddHyperlinkToWebNavbar** присвоено **значение True** или свойство **WebPageOptions.IncludePageOnNewWebNavigationBars** задано значение **True** для добавления или обновления в рамках указанного **WebNavigationBarSet**. Изменение этого параметра не изменяет количество элементов в панели, просто определяет, будут ли добавлены новые страницы. Установка этого значения в **значение False,** возможность разработки панелей навигации для определенных страниц на веб-сайте, не содержащих все доступные ссылки в публикации.

Значение по умолчанию — **True**. 


## <a name="example"></a>Пример

В следующем примере добавляется новый панель навигации задать в активный документ с помощью кнопок стиля текста и автоматическое обновление задано значение **False** , чтобы не будут добавлены ссылки на страницы или новые страницы автоматически обновляется на панели навигации панель навигации добавляется к первой страницы публикации.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newBar") 
With objWebNav 
 .AutoUpdate = False 
 .ButtonStyle = pbnbButtonStyleText 
End With 
ActiveDocument.Pages(1).Shapes.AddWebNavigationBar _ 
 Name:="newBar", Left:=10, Top:=10 

```


