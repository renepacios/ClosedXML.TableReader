using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.TableReader.Test.Fixtures;
using ClosedXML.TableReader.Test.Models;
using FluentAssertions;
using Xunit;

namespace ClosedXML.TableReader.Test.ReadDiferentTypes
{
    public class ReadWithNullableTypes : IClassFixture<CreateExcelConnectionFixture>
    {
        private readonly XLWorkbook _wb;
        public ReadWithNullableTypes(CreateExcelConnectionFixture connectionFixture)
        {
            _wb = connectionFixture.Init("TableWithAllTypes");
        }

        [Fact]
        public void ReadTable_Work_As_Should()
        {
            IEnumerable<AllTypes> data = _wb.ReadTable<AllTypes>(1);
            data.Should()
                .NotBeNull()
                .And
                .HaveCount(5);
        }
    }
}



