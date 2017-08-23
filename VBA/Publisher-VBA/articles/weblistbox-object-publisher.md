---
title: "Объект WebListBox (издатель)"
keywords: vbapb10.chm4128767
f1_keywords: vbapb10.chm4128767
ms.prod: publisher
api_name: Publisher.WebListBox
ms.assetid: 0ba881f8-95cf-c536-7fa8-05714348577d
ms.date: 06/08/2017
ms.openlocfilehash: 0960efeb0c421daf1f91fcac574d39216d4f970a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistbox-object-publisher"></a>Объект WebListBox (издатель)

Представляет элемент управления списка Web. Объект **WebListBox** является членом объекта **Shape** .
 


## <a name="example"></a>Пример

Используйте метод **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** для создания нового списка Web. Используйте свойство **[WebListBox](shape-weblistbox-property-publisher.md)** для доступа к фигурой Web поля списка элемента управления. Используйте метод **[AddItem](weblistboxitems-additem-method-publisher.md)** объекта **[WebListBoxItems](weblistboxitems-object-publisher.md)** для добавления элементов в поле со списком Web. В этом примере создается новый список Web и добавляет несколько элементов. Обратите внимание, что при создании, элемент управления списка Web содержит три элемента по умолчанию. В этом примере включает в себя процедуры, в котором удаляются поля элементов списка по умолчанию, прежде чем добавлять новые элементы.
 

 

 

 

 **Примечание**  При создании поля со списком Web начальной ширина — 300 точек. Тем не менее Microsoft Publisher автоматически изменяет этот ширины на основе ширины элементов в списке.
 




```
Sub CreateWebListBox() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72).WebListBox 
 .MultiSelect = msoFalse 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
 End With 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](weblistbox-application-property-publisher.md)|
|[ListBoxItems](weblistbox-listboxitems-property-publisher.md)|
|[MultiSelect](weblistbox-multiselect-property-publisher.md)|
|[Родительский раздел](weblistbox-parent-property-publisher.md)|
|[ReturnDataLabel](weblistbox-returndatalabel-property-publisher.md)|

