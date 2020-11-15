using System;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using FluentAssertions;
using Net.Core.Sample.Models;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.TableReader.Test.Models;
using Xunit;

namespace ClosedXML.TableReader.Test
{
    public class ReadUsingConventions : IClassFixture<CreateExcelConnectionFixture>
    {

        private readonly XLWorkbook _wb;

        public ReadUsingConventions(CreateExcelConnectionFixture connectionFixture)
        {


            _wb = connectionFixture.Init("TableWithHeaders");

        }

        [Fact]
        public void ReadTable_Work_As_Should()
        {
            IEnumerable<SimpleTable> data = _wb.ReadTable<SimpleTable>(1);

            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);

        }

        [Fact]
        public void ReadTable_Works_In_Row_Order()
        {
            IEnumerable<SimpleTable> data = _wb.ReadTable<SimpleTable>(1);

            var dataList = data.ToList();

            dataList.Should()
                .NotBeNull()
                .And
                .BeInAscendingOrder(s=>s.Number);


            dataList.Should()
                .NotBeNull()
                .And
                .BeInAscendingOrder(s => s.ADate);
        }



        [Fact]
        public void ReadTable_Works_With_Header_Mapping()
        {
            IEnumerable<TableWithHeaders> data = _wb.ReadTable<TableWithHeaders>(1);

            var dataList = data?.ToList();

            dataList.Should()
                .NotBeNull()
                .And
                .HaveCount(5);

          
            dataList.Should()
                .NotBeNull()
                .And
                .BeInAscendingOrder(s => s.Birthday);


            dataList.FirstOrDefault()?.Birthday.Should().Be(new DateTime(1980, 5, 10));

        }
    }

}

