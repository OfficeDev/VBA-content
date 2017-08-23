---
title: "Свойство Application.FileDialog (издатель)"
keywords: vbapb10.chm131089
f1_keywords: vbapb10.chm131089
ms.prod: publisher
api_name: Publisher.Application.FileDialog
ms.assetid: 65d73a9d-be4c-d809-d10d-468181ef9eb0
ms.date: 06/08/2017
ms.openlocfilehash: a814b13e66e5f2d493281f91f08a10af203c8ae3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationfiledialog-property-publisher"></a>Свойство Application.FileDialog (издатель)

Возвращает объект **классов FileDialog** , представляющий один экземпляр диалогового окна файла.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Классов FileDialog** ( **_Тип_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **MsoFileDialogType**| Тип диалогового окна.|

### <a name="return-value"></a>Возвращаемое значение

Классов FileDialog


## <a name="remarks"></a>Заметки

Тип параметров может иметь одно из ** [MsoFileDialogType](http://msdn.microsoft.com/library/ee445a67-1193-f446-4bd2-963c07fba5ae%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере отображается диалоговое окно " **Сохранить как** " и сохраняет имя файла, указанного пользователем.


```vb
Sub ShowSaveAsDialog() 
 Dim dlgSaveAs As FileDialog 
 Dim strFile As String 
 
 Set dlgSaveAs = Application.FileDialog( _ 
 Type:=msoFileDialogSaveAs) 
 dlgSaveAs.Show 
 strFile = dlgSaveAs.SelectedItems(1) 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

