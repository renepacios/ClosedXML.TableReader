using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ClosedXML.Excel;
using ClosedXML.TableReader.Model;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test
{
    public class ReadUsingConverters : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadUsingConverters(CreateExcelConnectionFixture connectionFixture)
        {
            _wb = connectionFixture.Init("TableWithHeaders");
        }

        [Fact]
        public void ReadTable_WithOutConvertersExpressions_Work_As_Should()
        {
            IEnumerable<TableWithHeadersAndBoleanConversion> data = _wb.ReadTable<TableWithHeadersAndBoleanConversion>(1);

            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);


            data.Select(s => s.Selected)
                .Should()
                .AllBeEquivalentTo(false);

        }

        [Fact]
        public void ReadTable_WithOutConvertersExpressions_And_NullableTypes_Work_As_Should()
        {
            IEnumerable<TableWithHeadersAndNullableBoleanConversion> data = _wb.ReadTable<TableWithHeadersAndNullableBoleanConversion>(1);

            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);


            data.Select(s => s.Selected)
                .Should()
                .AllBeEquivalentTo<bool?>(null);

        }

        [Fact]
        public void ReadTable_WithConvertersExpressions_Work_As_Should()
        {
            Expression<Func<string, bool>> fSelectToBool = c => c == "x";

            IEnumerable<TableWithHeadersAndBoleanConversion> data = _wb
                .ReadTable<TableWithHeadersAndBoleanConversion>(1,
                    new ReadOptions()
                    {
                        TitlesInFirstRow = true,
                        Converters = new Dictionary<string, LambdaExpression>()
                        {
                            {nameof(TableWithHeadersAndBoleanConversion.Selected),fSelectToBool },
                        }
                    });


            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);



            data
                .Where(w => new[] { 1, 3, 5 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(3)
                .And
                .AllBeEquivalentTo(true);

            data
                .Where(w => new[] { 2, 4 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(2)
                .And
                .AllBeEquivalentTo(false);

        }

        [Fact]
        public void ReadTable_WithConvertersExpressions_And_NullableTypes_Work_As_Should()
        {
            Expression<Func<string, bool?>> fSelectToBool = c => c == "x" ? true : (bool?)null;

            IEnumerable<TableWithHeadersAndNullableBoleanConversion> data = _wb
                .ReadTable<TableWithHeadersAndNullableBoleanConversion>(1,
                    new ReadOptions()
                    {
                        TitlesInFirstRow = true,
                        Converters = new Dictionary<string, LambdaExpression>()
                        {
                            {nameof(TableWithHeadersAndBoleanConversion.Selected),fSelectToBool },
                        }
                    });


            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);



            data
                .Where(w => new[] { 1, 3, 5 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(3)
                .And
                .AllBeEquivalentTo<bool?>(true);

            data
                .Where(w => new[] { 2, 4 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(2)
                .And
                .AllBeEquivalentTo<bool?>(null);

        }



        [Fact]
        public void ReadTable_WithConvertersExpressions_And_NullableTypes_Implicit_Conversion_Works_As_Should()
        {
            Expression<Func<string, bool>> fSelectToBool = c => 
                c == "x";

            IEnumerable<TableWithHeadersAndNullableBoleanConversion> data = _wb
                .ReadTable<TableWithHeadersAndNullableBoleanConversion>(1,
                    new ReadOptions()
                    {
                        TitlesInFirstRow = true,
                        Converters = new Dictionary<string, LambdaExpression>()
                        {
                            {nameof(TableWithHeadersAndBoleanConversion.Selected),fSelectToBool },
                        }
                    });


            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);



            data
                .Where(w => new[] { 1, 3, 5 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(3)
                .And
                .AllBeEquivalentTo<bool?>(true);

            data
                .Where(w => new[] { 2, 4 }.Contains(w.Number))
                .Select(s => s.Selected)
                .Should()
                .HaveCount(2)
                .And
                .AllBeEquivalentTo<bool?>(null);

        }


       
    }
}
