using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test.ReadDiferentTypes
{
    public class ReadChart : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadChart(CreateExcelConnectionFixture connectionFixture)
        {
            _wb = connectionFixture.Init("TableWithAllTypes");
        }

        [Fact]
        public void ReadDocument_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);
            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);
        }



        [Fact]
        public void ReadNull_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Null");

            d.Should().NotBeNull();
            d.Char.Should().Be('\0');

        }


        [Fact]
        public void ReadNull_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);

            var d = data.First(w => w.Description == "Null");

            d.Should().NotBeNull();
            d.Char.Should().Be(null);

        }

        [Fact]
        public void ReadLowerCase_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.Char.Should().Be('a');

        }

        [Fact]
        public void ReadLowerCasePositive_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.Char.Should().Be('a');
        }



        [Fact]
        public void ReadUpperCase_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.Char.Should().Be('C');

        }

        [Fact]
        public void ReadUpperCase_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.Char.Should().Be('C');
        }



        [Fact]
        public void ReadCharNumber_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.Char.Should().Be('1');

        }

        [Fact]
        public void ReadCharNumber_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.Char.Should().Be('1');
        }



        [Fact]
        public void ReadUnicodeChar_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.Char.Should().Be('@');

        }

        [Fact]
        public void ReadUnicodeChar_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.Char.Should().Be('@');
        }














    }
}