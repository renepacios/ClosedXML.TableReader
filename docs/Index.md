# ClosedXML.TableReader

ClosedXml is an amazing .NET library to manage Excel files based on [OpenXML Standard](https://es.wikipedia.org/wiki/Office_Open_XML). This component drastically reduces the complexity of working with [Open XML SDK](https://www.microsoft.com/en-us/download/details.aspx?id=30425) 

**ClosedXML.TableReader** extends ClosedXML to read Excel tables in typed way. You can read all or some columns, apply transformation rules for each field, refer column by title or name using DataAnotations.

  <br>

## Read Table (using conventions)

By default, library read table from first column and row used and assume that columns titles are in first row. 
For this sample we have next Excel table:

![](./images/SimpleTable.png?raw=true)

We define this VO class:
```csharp 

  public class SimpleTable
    {
        public int Number { get; set; } 

        public string Name { get; set; }

        public DateTime ADate { get; set; }

    }
```
And we could read table
```csharp 
  byte[] file = System.IO.File.ReadAllBytes([ExcelPath])

            var wb = new ClosedXML.Excel.XLWorkbook(new MemoryStream(file));

            IEnumerable<SimpleTable> data = wb.ReadTable<SimpleTable>(1); // indicate excel sheet
```
>We have to indicate number of sheet where is the table. (First Excel document sheet is number 1)  

 
  
  <br>  <br>

## Data Anotations

Title columns used to more complex that *Number, Name or ADate* and might not match with our properties name conventions. In this case can use DataAnnotations to pair column with object field. 

For example, next Excel shows a new column with title *Birthday date*

![](./images/DataAnotations.png?raw=true)

We pair with field using standard DisplayName Annotation 
```csharp 

  public class SimpleTable
    {
        public int Number { get; set; } 

        public string Name { get; set; }

        [DisplayName("Birthday date")]
        public DateTime ADate { get; set; }

    }
```
  <br>

#### Use anotations to read table without title
We can used DataAnnotations to refer columns without titles, in this case we have to use builtin custom annotation **ColumnName**

For Example, for this Excel Table
![](./images/SimpleTableWithOutTitles.png?raw=true)

We have to modify VO class like that
```csharp 

  public class SimpleTable
    {
        [ColumnName("B")]
        public int Number { get; set; } 

        [ColumnName("C")]
        public string Name { get; set; }

        [ColumnName("D")]
        public DateTime ADate { get; set; }

    }
```
  <br>

**_Don't forget say library that first row isn't a title in ReadOptions_**

```csharp
  var data = wb.ReadTable<SimpleTable>(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = false
                });
```

## Data transformations

Optionally, we can make transformations over data before populate in memory object. We can indicate an *Expression (per field)* for transform any data from Excel cell to adapt to target type


![](./images/Transformations.png?raw=true)

For example, we have column selected in Excel document with a ‘x’ mark and we want transform this data in Boolean value.
```csharp

            //Example with trasnformations
            Expression<Func<string, DateTime>> fConvertData = _ => DateTime.Now.AddDays(10);
            Expression<Func<string, bool>> fSelectToBool = c => c == "x";

            var ls = wb.ReadTable<Models.TableWithHeadersAndBoleanConversion>(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = true,
                    Converters = new Dictionary<string, LambdaExpression>()
                    {
                        {nameof(SimpleTable.ADate),fConvertData },
                        {nameof(SimpleTable.Selected),fSelectToBool },
                    }
                });
```

As show, we can transform any value from source. 
- Expression are defined in **ReadOptions.Converters** dictionary. 
- Dictionary entry key should be object field name, obviusly we can use *nameof*  for that
