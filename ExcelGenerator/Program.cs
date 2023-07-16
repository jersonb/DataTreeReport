using var workbook = new XLWorkbook();

var worksheet = workbook.Worksheets.Add("Sample Sheet");

var variables = GetVariablesMock();

Print(worksheet, variables);

var filePath = @"./test.xlsx";

File.Delete(filePath);

workbook.SaveAs(filePath);

Process.Start(@"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE", filePath);

static List<Variable> GetVariablesMock()
{
    var variable1 = new Variable
    {
        Name = "variable1",
        Level = VariableLevel.LEVEL1,
        Children = new List<Variable>
        {
            new ()
            {
                Name = "variable11",
                Level = VariableLevel.LEVEL2,
                Children = new List<Variable>
                {
                    new ()
                    {
                        Name = "variable111",
                        Level = VariableLevel.LEVEL3,
                    },
                    new ()
                    {
                        Name = "variable112",
                        Level = VariableLevel.LEVEL3,
                    },
                    new ()
                    {
                        Name = "variable113",
                        Level = VariableLevel.LEVEL3,
                    },
                }
            },
            new ()
            {
                Name = "variable12",
                Level = VariableLevel.LEVEL2,
                Children = new List<Variable>
                {
                    new ()
                    {
                        Name = "variable121",
                        Level = VariableLevel.LEVEL3,
                    },
                    new ()
                    {
                        Name = "variable122",
                        Level = VariableLevel.LEVEL3,
                    },
                }
            },
        }
    };

    var variable2 = new Variable
    {
        Name = "variable2",
        Level = VariableLevel.LEVEL1,
        Children = new List<Variable>
        {
            new ()
            {
                Name = "variable21",
                Level = VariableLevel.LEVEL2,
                Children = new List<Variable>
                {
                    new ()
                    {
                        Name = "variable211",
                        Level = VariableLevel.LEVEL3,
                    },
                    new ()
                    {
                        Name = "variable212",
                        Level = VariableLevel.LEVEL3,
                    }
                }
            },
            new ()
            {
                Name = "variable22",
                Level = VariableLevel.LEVEL2,
                Children = new List<Variable>
                {
                    new ()
                    {
                        Name = "variable221",
                        Level = VariableLevel.LEVEL3,
                    },
                }
            },
        }
    };

    return new()
    {
        variable1,
        variable2,
    };
}

static void Print(IXLWorksheet worksheet, List<Variable> variables)
{
    variables.ForEach(variable =>
    {
        var column = variable.Level switch
        {
            VariableLevel.LEVEL1 => Column.A,
            VariableLevel.LEVEL2 => Column.B,
            VariableLevel.LEVEL3 => Column.C,
            VariableLevel.LEVEL4 => Column.C,
            _ => Column.ERROR,
        };

        var lastCell = worksheet.Column(column).LastCellUsed(XLCellsUsedOptions.MergedRanges | XLCellsUsedOptions.AllContents);
        var lastRow = lastCell != null ? lastCell.Address.RowNumber : 0;

        lastRow++;

        if (variable.Children.Any())
        {
            worksheet
            .Range(lastRow, column, (lastRow + variable.LeafCount) - 1, column)
            .Merge()
            .SetValue(variable.Name)
            .Style
            .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            Print(worksheet, variable.Children);
        }
        else
        {
            worksheet
            .Cell(lastRow, column)
            .SetValue(variable.Name);
        }
    });
}