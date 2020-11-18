using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test.ReadDiferentTypes
{
    public class ReadDate : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadDate(CreateExcelConnectionFixture connectionFixture)
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
            d.Date.Should().Be(DateTime.MinValue);

        }


        [Fact]
        public void ReadNull_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);

            var d = data.First(w => w.Description == "Null");

            d.Should().NotBeNull();
            d.Date.Should().BeNull();
            d.Date.Should().Be((DateTime?)null);

        }

        [Fact]
        public void ReadPositive_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            //10/05/1980  12:30:00
            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 10,12,30,00));

        }

        [Fact]
        public void ReadPositive_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 10, 12, 30, 00));
        }



        [Fact]
        public void ReadBigger_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 11,14, 33, 00));

        }

        [Fact]
        public void ReadBigger_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 11, 14, 33, 00));
        }



        [Fact]
        public void ReadNegative_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 12, 4, 24, 22));

        }

        [Fact]
        public void ReadNegative_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.Date.Should().Be(new DateTime(1980, 05, 12, 4, 24, 22));
        }



        [Fact]
        public void ReadMin_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.Date.Should().Be(DateTime.MinValue);

        }

        [Fact]
        public void ReadMin_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.Date.Should().Be(DateTime.MinValue);
        }














    }
}