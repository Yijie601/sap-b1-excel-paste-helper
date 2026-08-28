using System.Text;

namespace SapB1ExcelHelper.Services;

public sealed record SupplierMappingEntry(string SupplierName, string SapCode);

public sealed class SupplierMappingService
{
    private readonly string _filePath;
    private Dictionary<string, string> _mappings = new(StringComparer.OrdinalIgnoreCase);

    public SupplierMappingService(string? filePath = null)
    {
        _filePath = filePath ?? AppPaths.SupplierMappingFile;
        Reload();
    }

    public int Count => _mappings.Count;

    public void Reload()
    {
        _mappings = LoadEntries(_filePath)
            .GroupBy(item => item.SupplierName.Trim(), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Key.Length > 0)
            .ToDictionary(
                group => group.Key,
                group => group.Last().SapCode.Trim(),
                StringComparer.OrdinalIgnoreCase);
    }

    public bool TryResolve(string supplierName, out string sapCode) =>
        _mappings.TryGetValue(supplierName.Trim(), out sapCode!);

    public IReadOnlyList<SupplierMappingEntry> GetAll() => _mappings
        .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
        .Select(pair => new SupplierMappingEntry(pair.Key, pair.Value))
        .ToArray();

    public void Save(IEnumerable<SupplierMappingEntry> entries)
    {
        var validEntries = entries
            .Select(entry => new SupplierMappingEntry(entry.SupplierName.Trim(), entry.SapCode.Trim()))
            .Where(entry => entry.SupplierName.Length > 0 && entry.SapCode.Length > 0)
            .GroupBy(entry => entry.SupplierName, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.Last())
            .OrderBy(entry => entry.SupplierName, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var builder = new StringBuilder("Supplier Name,SAP Code\r\n");
        foreach (var entry in validEntries)
        {
            builder.Append(EscapeCsv(entry.SupplierName))
                .Append(',')
                .Append(EscapeCsv(entry.SapCode))
                .Append("\r\n");
        }

        File.WriteAllText(_filePath, builder.ToString(), new UTF8Encoding(false));
        Reload();
    }

    public IReadOnlyList<SupplierMappingEntry> Import(string filePath) => LoadEntries(filePath);

    private static IReadOnlyList<SupplierMappingEntry> LoadEntries(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return Array.Empty<SupplierMappingEntry>();
        }

        var result = new List<SupplierMappingEntry>();
        var lines = File.ReadAllLines(filePath);
        foreach (var line in lines.Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var fields = ParseCsvLine(line);
            if (fields.Count >= 2)
            {
                result.Add(new SupplierMappingEntry(fields[0].Trim(), fields[1].Trim()));
            }
        }

        return result;
    }

    private static IReadOnlyList<string> ParseCsvLine(string line)
    {
        var fields = new List<string>();
        var field = new StringBuilder();
        var quoted = false;

        for (var index = 0; index < line.Length; index++)
        {
            var character = line[index];
            if (character == '"')
            {
                if (quoted && index + 1 < line.Length && line[index + 1] == '"')
                {
                    field.Append('"');
                    index++;
                }
                else
                {
                    quoted = !quoted;
                }
            }
            else if (character == ',' && !quoted)
            {
                fields.Add(field.ToString());
                field.Clear();
            }
            else
            {
                field.Append(character);
            }
        }

        fields.Add(field.ToString());
        return fields;
    }

    private static string EscapeCsv(string value)
    {
        if (!value.Contains(',') && !value.Contains('"') && !value.Contains('\r') && !value.Contains('\n'))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}

