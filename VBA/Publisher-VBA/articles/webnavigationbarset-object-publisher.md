---
title: "Объект WebNavigationBarSet (издатель)"
keywords: vbapb10.chm8585215
f1_keywords: vbapb10.chm8585215
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet
ms.assetid: 03b31cc1-5b24-1a16-710c-73755298066e
ms.date: 06/08/2017
ms.openlocfilehash: e5121d2ffc0a3c35201b6a80efb8458a2ddef469
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarset-object-publisher"></a>Объект WebNavigationBarSet (издатель)

Представляет панель навигации для текущего документа. Объект **WebNavigationBarSet** , является участником коллекцию **WebNavigationBarSets** , которая включает в себя все веб-сайта, задает панели навигации в текущем документе.
 


## <a name="example"></a>Пример

С помощью **WebNavigationBarSet**. **AddToEveryPage** (Слева, сверху, [ширина]), где слева — это положение левого края фигуры в начало — это положение верхней границы фигуры и ширина — ширину фигуры, представляющий набор панель навигации Web, для добавления указанного панель навигации на каждой странице документа. В следующем примере добавляется первого панель навигации задайте каждые страницу, которая содержит свойство **AddHyperlinkToWebNavbar** задано значение **True** , при добавлении на страницу или свойство **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** задано значение **True**.
 

 

```
Dim objWebNavBarSet as WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1) 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```

С помощью **WebNavigationBarSet**. **DeleteSetAndInstances** для удаления Web навигационной панели набора и каждый экземпляр из документа. Следующий пример удаляет все экземпляры каждого объекта **WebNavigationBarSet** в коллекции **WebNavigationBarSets** .
 

 



```
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```

Существует три свойства, касающиеся горизонтальный панелей навигации Web. С помощью **WebNavigationBarSet**. Задайте **IsHorizontal** для определения ориентации на панели навигации. Метод **ChangeOrientation** используется для установки ориентации набор панели навигации веб. Если на **горизонтальную**ориентацию, свойства **HorizontalAlignment** и **HorizontalButtonCount** можно задать. В следующем примере добавляет первой панели навигации в коллекции **WebNavigationBarSets** активного документа для каждой страницы, имеет свойство **AddHyperlinkToWebNavbar** задано значение **True** или свойство **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** задано значение **True**и затем задает стиль кнопки для **малого**. Для определения, является ли набор панель навигации горизонтальной выполняется проверка. Если он не установлен, вызывается метод **ChangeOrientation** , ориентации задано значение **Горизонтальная**. После ориентирована на панели навигации по горизонтали, count горизонтальной кнопки задано значение **3** , горизонтальное выравнивание кнопок задано значение **слева**.
 

 



```
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignLeft 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AddToEveryPage](webnavigationbarset-addtoeverypage-method-publisher.md)|
|[ChangeOrientation](webnavigationbarset-changeorientation-method-publisher.md)|
|[DeleteSetAndInstances](webnavigationbarset-deletesetandinstances-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webnavigationbarset-application-property-publisher.md)|
|[Автоматическое обновление](webnavigationbarset-autoupdate-property-publisher.md)|
|[ButtonStyle](webnavigationbarset-buttonstyle-property-publisher.md)|
|[Разработка](webnavigationbarset-design-property-publisher.md)|
|[HorizontalAlignment](webnavigationbarset-horizontalalignment-property-publisher.md)|
|[HorizontalButtonCount](webnavigationbarset-horizontalbuttoncount-property-publisher.md)|
|[IsHorizontal](webnavigationbarset-ishorizontal-property-publisher.md)|
|[Ссылки](webnavigationbarset-links-property-publisher.md)|
|[Name](webnavigationbarset-name-property-publisher.md)|
|[Родительский раздел](webnavigationbarset-parent-property-publisher.md)|
|[ShowSelected](webnavigationbarset-showselected-property-publisher.md)|

