# ExcelService
- A Fluent Interface for OpenXML with Dynamic Object XLSX Creation and Modification. 
- Includes Styling based on query or based on location, using base Open XML properties
- Fast Quick Table Creation and Modification Without DataTable use

# Installation

## Dotnet CLI
`
dotnet add package FluentExcelService --version 1.2.0
`
## Package Manager  
`
NuGet\Install-Package FluentExcelService -Version 1.2.0
`
# Usage

## Locally Saving a File 
```csharp
Excel excel = new Excel();

excel.GenerateNewWorkBook(new List<TestClass>()
{
    new TestClass("a", "b", "c", "d", "e", "f", "g", 1, new DateTime(1999, 12, 08)),
    new TestClass("h", "i", "j", "k", "l", "m", "m", 2, new DateTime(1998, 11, 07)),
    new TestClass("n", "l", "o", "p", "q", "r", "s", 2, new DateTime(1998, 11, 07)),
},
new List<List<Style>>()
{ 
    new List<Style>()
    { 
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
    },
    new List<Style>()
    {
        new Style(Enums.Font.Calibri, Color.Red, 8),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        new Style(Enums.Font.Calibri, Color.Red, 8),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
    },
    new List<Style>()
    {
        new Style(Enums.Font.Calibri, Color.Pink, 8),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        new Style(Enums.Font.Arial, Color.Green, 20),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
    },
    new List<Style>()
    {
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
        new Style(Enums.Font.Arial, Color.Red, 25),
        new Style(Enums.Font.Arial, Color.Yellow, 48),
        Style.Empty(),
        Style.Empty(),
        Style.Empty(),
    },
},
"TestWorkbook",
"TestSheet");

excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "a", new Style(Enums.Font.Arial, Color.Green, 65));

excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "n", new Style(Enums.Font.Arial, Color.Red, 25));

excel["TestWorkbook"]?.StyleCellWhere<TestClass>(x => x.Column2 == "b", new Style(Enums.Font.Calibri, Color.Aqua, 30));

excel.SaveExcelFileFromWorkbook("../../../test.xlsx", excel["TestWorkbook"] ?? throw new NullReferenceException("Invalid Container"));
```

## Getting a Workbook From a File

```csharp

Excel excel = new Excel();

Workbook workbook = excel.GetWorkbookFromExcelFile("../../../test.xlsx");


```

## Dependency Injections
```csharp
//Add Interface and Class as normal... for azure as example:
services.AddScoped<IExcel, Excel>();

//-----------------------------

public class MyDependencyInjectableClass
{
    private readonly _excel;
    public MyDependencyInjectableClass(IExcel excel)
    {
        _excel = excel;
    }

    public void Foo()
    {
        _excel..... methods
    }
}
```
## Saving Stream to a File
```csharp

_excel.GenerateNewWorkBook(Workbook.GetWorkbookFromDataSet(myEnumerableOfObjects, null, "A WorkSheet Name", "A sheet"));
_excel.GenerateNewWorkBook(Workbook.GetWorkbookFromDataSet(myOtherEnumerableOfObjects, null, "A Second WorkSheet Name", "A sheet"));

_excel["A WorkSheet Name"]!.StyleRowWhere(x => x.MyParameter == "a really cool parameter")

_excel.CombineWorkbooks(_excel["A WorkSheet Name"]!, _excel["A Second WorkSheet Name"]!);
Stream myOtherStream = _excel.GetExcelFromWorkBook(_excel["DARTContainersNotTrackingInCavi"]!);
```
