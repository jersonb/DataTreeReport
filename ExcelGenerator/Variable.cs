namespace ExcelGenerator;

public class Variable
{
    public Guid Uuid { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty!;
    public VariableLevel Level { get; set; }
    public List<Variable> Children { get; set; } = new List<Variable>();

    public int LeafCount => GetLeafCount(this);

    private int GetLeafCount(Variable variable)
    {
        var variables = variable.Children;
        var sum = variables.Sum(t => t.Children.Count);

        if (sum == 0)
            return variables.Count;

        return variables.Sum(GetLeafCount);
    }
}
