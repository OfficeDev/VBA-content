---
title: "Объект WebNavigationBarSets (издатель)"
keywords: vbapb10.chm8519679
f1_keywords: vbapb10.chm8519679
ms.prod: publisher
api_name: Publisher.WebNavigationBarSets
ms.assetid: 0c4f62c7-b7b2-a7bc-60f8-8097fe99fe58
ms.date: 06/08/2017
ms.openlocfilehash: b3985b719f64716c5d2d78baaea4bf1b0fcf68cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsets-object-publisher"></a>Объект WebNavigationBarSets (издатель)

Коллекция всех объектов **WebNavigationBarSet** в текущем документе. Каждый **WebNavigationBarSet** представляет панель навигации, набор, состоящий из гиперссылки.
 


## <a name="remarks"></a>Заметки

По умолчанию существует два объекта **WebNavigationBarSet** на каждом веб-странице мастера; один только текст, а другое — вертикальной. Эти объекты соответствуют разработки мастер вне зависимости от того, является ли используемый панели навигации на странице.
 

 

## <a name="example"></a>Пример

Возвращает объект **WebNavigationBarSet** , используйте свойство **WebNavigationBarSets** текущего документа. В следующем примере задается переменная объекта в коллекцию **WebNavigationBarSets** активных документов.
 

 

```
Dim objWebNavBarSets As WebNavigationBarSets 
Set objWebNavBarSets = ActiveDocument.WebNavigationBarSets
```

Используйте **WebNavigationBarSets** **. Элемент** (индекс), где индекс — номер индекса, чтобы возвратить объект **WebNavigationBarSet** из коллекции. В следующем примере возвращается первый панель навигации задайте из коллекции **WebNavigationBarSets** .
 

 



```
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.Item(1)
```

Предыдущий пример также может быть выполнено с помощью **WebNavigationBarSets** (индекс), где индекс — номер индекса, чтобы возвратить объект **WebNavigationBarSet** .
 

 



```
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1)
```

Предыдущий пример также может быть выполнено с помощью **WebNavigationBarSets** (индекс) где индекса — это строка, указывающая имя панель навигации задать для возвращения.
 

 



```
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets("WebNavBarSet1")
```

Используйте **WebNavigationBarSets** **. Число** возвращает число веб-переходе задает панели в коллекции. В этом примере отображается число веб-панель инструментов задает текущего документа.
 

 



```
MsgBox ActiveDocument.WebNavigationBarSets.Count 

```

Используйте **WebNavigationBarSets** **. AddToEveryPage** (слева, сверху, [ширина]), где слева — расстояние от левого края страницы на левой панели навигации, находится на форме расстояние в верхней части страницы по верхнему краю панели навигации и ширина — ширину панели navigaion для добавления на панели навигации, указанный для каждой страницы. В следующем примере добавляется на панели навигации, с именем «WebNavBar1» на все страницы публикации.
 

 



```
ActiveDocument.WebNavigationBarSets.Item _ 
 ("WebNavBarSet1").AddToEveryPage _ 
 Left:=50, Top:=25
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AddSet](webnavigationbarsets-addset-method-publisher.md)|
|[Элемент](webnavigationbarsets-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webnavigationbarsets-application-property-publisher.md)|
|[Count](webnavigationbarsets-count-property-publisher.md)|
|[Родительский раздел](webnavigationbarsets-parent-property-publisher.md)|

