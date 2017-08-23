---
title: "Свойство WebCommandButton.DataFileName (издатель)"
keywords: vbapb10.chm3932165
f1_keywords: vbapb10.chm3932165
ms.prod: publisher
api_name: Publisher.WebCommandButton.DataFileName
ms.assetid: 5fd2bac7-7067-4833-4b34-26897c39ea58
ms.date: 06/08/2017
ms.openlocfilehash: eb0f9eb62890be6b4201331f9b221d16570772e5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttondatafilename-property-publisher"></a>Свойство WebCommandButton.DataFileName (издатель)

Возвращает или задает **строку** , представляющую имя файла для сохранения данных из веб-форму. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя_файла_данных**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере задается Microsoft Publisher процесс данных веб-форм путем сохранения файла с разделителями-запятыми на одном веб-сервере, как форма будет сохранена.


```vb
Sub WebDataFile() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalSaveOnServer 
 .DataFileFormat = pbSubmitDataFormatCSV 
 .DataFileName = "WebFormData.txt" 
 End With 
End Sub
```


