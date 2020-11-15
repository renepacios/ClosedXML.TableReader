using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test.ReadDiferentTypes
{
    

    public class ReadString : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadString(CreateExcelConnectionFixture connectionFixture)
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
            d.String.Should().Be(string.Empty);

        }


        [Fact]
        public void ReadNull_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);

            var d = data.First(w => w.Description == "Null");

            d.Should().NotBeNull();
            d.String.Should().Be(string.Empty);

        }

        [Fact]
        public void ReadLowerCase_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.String.Should().Be("jhon doe");

        }

        [Fact]
        public void ReadLowerCasePositive_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
            d.String.Should().Be("jhon doe");
        }



        [Fact]
        public void ReadUpperCase_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.String.Should().Be("PETTER PARKER");

        }

        [Fact]
        public void ReadUpperCase_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
            d.String.Should().Be("PETTER PARKER");
        }



        [Fact]
        public void ReadNumber_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.String.Should().Be("1");

        }

        [Fact]
        public void ReadNumber_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
            d.String.Should().Be("1");
        }



        [Fact]
        public void ReadSpace_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.String.Should().Be(" ");

        }

        [Fact]
        public void ReadSpace_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
            d.String.Should().Be(" ");
        }














    }


}