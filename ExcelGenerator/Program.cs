using System.Diagnostics;
using ClosedXML.Excel;

using var workbook = new XLWorkbook();

var worksheet = workbook.Worksheets.Add("Sample Sheet");

var variables = GetVariablesMock();
variables.ForEach(SetParent);

var atualRow = 0;
variables.ForEach(variable =>
{
    var rangeUsed = SetValueRangeVertical(
        worksheet,
        1,
        atualRow + 1,
        LeafCount(variable),
        variable.Name);

    atualRow = rangeUsed.FirstCell().Address.RowNumber;

    variable.Children.ForEach(c1 =>
    {
        if (c1.Children.Any())
        {
            var rangeUsed = SetValueRangeVertical(
            worksheet,
            2,
            atualRow,
            LeafCount(c1),
            c1.Name);

            atualRow = rangeUsed.LastCell().Address.RowNumber + 1;
        }
        else
        {
            worksheet.Cell(atualRow, 2).Value = c1.Name;
            atualRow++;
        }
    });

    atualRow = rangeUsed.LastCell().Address.RowNumber;
});

var filePath = @"C:/dev/ExcelGenerator/ExcelGenerator/test.xlsx";

File.Delete(filePath);

workbook.SaveAs(filePath);

Process.Start(@"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE", filePath);

static int LeafCount(Variable test)
{
    var tests = test.Children;
    var sum = tests.Sum(t => t.Children.Count);

    if (sum == 0)
        return tests.Count;

    return tests.Sum(LeafCount);
}

static IXLRange SetValueRangeVertical(IXLWorksheet worksheet, int column, int firstRow, int rangeSize, string value)
{
    var range = worksheet.Range(firstRow, column, firstRow - 1 + rangeSize, column);
    range.Merge();
    range.SetValue(value);

    var alignment = range.Style.Alignment;
    alignment.SetTextRotation(255);
    alignment.SetVertical(XLAlignmentVerticalValues.Center);
    return range;
}

static List<Variable> GetVariablesMock()
{
    var variable1 = new Variable
    {
        Name = "variable1",
        Children = new List<Variable>
        {
            new (){ Name = "variable11" },
            new ()
            {
                Name = "variable12",
                Children =new List<Variable>
                {
                    new (){ Name = "variable121" },
                    new (){ Name = "variable122" }
                }
            },
        }
    };

    var variable2 = new Variable
    {
        Name = "variable2",
        Children = new List<Variable>
        {
            new (){ Name = "variable21" },
            new (){ Name = "variable22" }
        }
    };

    var variable3 = new Variable
    {
        Name = "variable3",
        Children = new List<Variable>
        {
            new ()
            {
                Name =  "variable31",
                Children = new ()
                {
                    new Variable
                    {
                        Name = "variable311",
                        Children = new List<Variable>
                        {
                            new (){ Name = "variable3111" },
                            new (){ Name = "variable3112" },
                            new (){ Name = "variable3113" },
                        }
                    }
                }
            }
        }
    };

    return new()
    {
        variable1,
        variable2,
        variable3
    };
}

static void SetParent(Variable variable)
{
    variable.Children.ForEach(v =>
    {
        v.Parent = variable;
        SetParent(v);
    });
}

public class Variable
{
    public Guid Uuid { get; set; } = Guid.NewGuid();
    public string Name { get; set; }
    public Variable Parent { get; set; }
    public List<Variable> Children { get; set; } = new List<Variable>();
}