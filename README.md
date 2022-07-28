# Flexcel
Fluent library for excel exporting

![](https://github.com/bannikovilea/Flexcel/workflows/.NET%20Core%20build%20and%20test/badge.svg)
![](https://img.shields.io/nuget/v/Flexcel.FluentExcelBuilder.svg?style=flat-square&label=nuget)

## Show me the code

Specific fields of model, field2 with custom title

```
// MyModel model;
ExcelFilesCreator.NewExcel()
                 .AddSheetAndSelect<MyModel>()
				 .AddColumn(m => m.Field1)
				 .AddColumn(m => m.Field2, "Another field")
				 .AddRow(model)
				 .ToFile(filePath);
```

All model fields

```
// MyModel model;
ExcelFilesCreator.NewExcel()
                 .AddSheetAndSelect<MyModel>()
				 .AddRow(model)
				 .ToFile(filePath);

```

Allows you export to:

- `.ToByteArray()`
- `.ToFile(filePath)`
- `.ToFileAsync(filePath)`
- `.ToMemoryStream()`

or get EpPlus package: `.GetPackage()`
