---
title: "Метод Pages.FindByPageID (издатель)"
keywords: vbapb10.chm458759
f1_keywords: vbapb10.chm458759
ms.prod: publisher
api_name: Publisher.Pages.FindByPageID
ms.assetid: 23ff5e69-33b1-e394-9d09-7199eae19fe9
ms.date: 06/08/2017
ms.openlocfilehash: 68fb025f238c6c5311d9ad882fca638972933014
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesfindbypageid-method-publisher"></a>Метод Pages.FindByPageID (издатель)

Возвращает объект **[страницы](page-object-publisher.md)** , представляющий страница с номером указанный идентификатор. Каждая страница автоматически назначается уникальный Идентификационный номер при его создании. Свойство **[PageID](page-pageid-property-publisher.md)** возвращает номер идентификатора страницы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindByPageID** ( **_PageID_**)

 переменная _expression_A, представляет собой объект- **страниц** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PageID|Обязательное свойство.| **Длинный**|Указывает идентификатор страницы, которую необходимо возвратить. Publisher присваивает этот номер при создании страниц.|

### <a name="return-value"></a>Возвращаемое значение

Page


## <a name="remarks"></a>Заметки

В отличие от свойство **[PageIndex](page-pageindex-property-publisher.md)** свойство **PageID** объекта **страницы** не будет изменяться при страницы, чтобы добавить или изменить порядок страниц в публикации. Таким образом с помощью метода **FindByPageID** с идентификатором страница может быть более надежный способ возврата определенного объекта **Page** из коллекции **[Pages](pages-object-publisher.md)** , чем при использовании метода **Item**с номером страницы.


## <a name="example"></a>Пример

В этом примере показано, как получить уникальный Идентификационный номер для объекта **Page** , а затем использовать этот номер для возврата объекта **страницы** из коллекции **страниц** и добавьте новый фигуры на страницу.


```vb
Sub FindPage() 
 Dim lngPageID As Long 
 
 'Get page ID 
 lngPageID = ActiveDocument.Pages.Add(Count:=1, After:=1).PageID 
 
 'Use page ID to add a new shape to the page 
 ActiveDocument.Pages.FindByPageID(PageID:=lngPageID) _ 
 .Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=200, Top:=72, Width:=50, Height:=50 
 
End Sub
```


