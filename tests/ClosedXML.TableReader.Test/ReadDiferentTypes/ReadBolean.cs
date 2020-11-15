﻿using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test.ReadDiferentTypes
{
    public class ReadBolean : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadBolean(CreateExcelConnectionFixture connectionFixture)
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
            d.BoleanString.Should().Be(false);

        }


        [Fact]
        public void ReadNull_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);

            var d = data.First(w => w.Description == "Null");

            d.Should().NotBeNull();
           d.BoleanString.Should().BeNull();
           d.BoleanString.Should().Be((bool?)null);

        }

        [Fact]
        public void ReadPositive_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(true);

        }

        [Fact]
        public void ReadPositive_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Positive");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(true);
        }



        [Fact]
        public void ReadBigger_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(true);

        }

        [Fact]
        public void ReadBigger_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Bigger");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(true);
        }



        [Fact]
        public void ReadNegative_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(false);

        }

        [Fact]
        public void ReadNegative_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Negative");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(false);
        }



        [Fact]
        public void ReadMin_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);

            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(false);

        }

        [Fact]
        public void ReadMin_With_Nullable_Types_Work_As_Should()
        {
            IEnumerable<AllTypesWithNulls> data = _wb.ReadTable<AllTypesWithNulls>(1);
            var d = data.First(w => w.Description == "Min");
            d.Should().NotBeNull();
           d.BoleanString.Should().Be(false);
        }














    }
}