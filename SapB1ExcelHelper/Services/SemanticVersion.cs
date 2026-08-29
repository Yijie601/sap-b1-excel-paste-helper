namespace SapB1ExcelHelper.Services;

public sealed class SemanticVersion : IComparable<SemanticVersion>
{
    private readonly string[] _preReleaseIdentifiers;

    private SemanticVersion(int major, int minor, int patch, string[] preReleaseIdentifiers)
    {
        Major = major;
        Minor = minor;
        Patch = patch;
        _preReleaseIdentifiers = preReleaseIdentifiers;
    }

    public int Major { get; }
    public int Minor { get; }
    public int Patch { get; }
    public bool IsPreRelease => _preReleaseIdentifiers.Length > 0;

    public static SemanticVersion Parse(string value)
    {
        if (!TryParse(value, out var version))
        {
            throw new FormatException($"Invalid semantic version: {value}");
        }

        return version!;
    }

    public static bool TryParse(string? value, out SemanticVersion? version)
    {
        version = null;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var normalized = value.Trim();
        if (normalized.StartsWith('v') || normalized.StartsWith('V'))
        {
            normalized = normalized[1..];
        }

        var buildIndex = normalized.IndexOf('+');
        if (buildIndex >= 0)
        {
            normalized = normalized[..buildIndex];
        }

        var preReleaseIndex = normalized.IndexOf('-');
        var core = preReleaseIndex >= 0 ? normalized[..preReleaseIndex] : normalized;
        var preRelease = preReleaseIndex >= 0 ? normalized[(preReleaseIndex + 1)..] : string.Empty;
        var numbers = core.Split('.');
        if (numbers.Length != 3 ||
            !int.TryParse(numbers[0], out var major) ||
            !int.TryParse(numbers[1], out var minor) ||
            !int.TryParse(numbers[2], out var patch) ||
            major < 0 || minor < 0 || patch < 0)
        {
            return false;
        }

        var identifiers = preRelease.Length == 0
            ? Array.Empty<string>()
            : preRelease.Split('.');
        if (identifiers.Any(identifier => identifier.Length == 0 ||
            identifier.Any(character => !char.IsLetterOrDigit(character) && character != '-')))
        {
            return false;
        }

        version = new SemanticVersion(major, minor, patch, identifiers);
        return true;
    }

    public int CompareTo(SemanticVersion? other)
    {
        if (other is null)
        {
            return 1;
        }

        var coreComparison = Major.CompareTo(other.Major);
        if (coreComparison == 0)
        {
            coreComparison = Minor.CompareTo(other.Minor);
        }

        if (coreComparison == 0)
        {
            coreComparison = Patch.CompareTo(other.Patch);
        }

        if (coreComparison != 0)
        {
            return coreComparison;
        }

        if (!IsPreRelease && !other.IsPreRelease)
        {
            return 0;
        }

        if (!IsPreRelease)
        {
            return 1;
        }

        if (!other.IsPreRelease)
        {
            return -1;
        }

        var count = Math.Min(_preReleaseIdentifiers.Length, other._preReleaseIdentifiers.Length);
        for (var index = 0; index < count; index++)
        {
            var left = _preReleaseIdentifiers[index];
            var right = other._preReleaseIdentifiers[index];
            var leftIsNumber = int.TryParse(left, out var leftNumber);
            var rightIsNumber = int.TryParse(right, out var rightNumber);

            int comparison;
            if (leftIsNumber && rightIsNumber)
            {
                comparison = leftNumber.CompareTo(rightNumber);
            }
            else if (leftIsNumber)
            {
                comparison = -1;
            }
            else if (rightIsNumber)
            {
                comparison = 1;
            }
            else
            {
                comparison = string.Compare(left, right, StringComparison.Ordinal);
            }

            if (comparison != 0)
            {
                return comparison;
            }
        }

        return _preReleaseIdentifiers.Length.CompareTo(other._preReleaseIdentifiers.Length);
    }

    public override string ToString()
    {
        var core = $"{Major}.{Minor}.{Patch}";
        return IsPreRelease ? $"{core}-{string.Join('.', _preReleaseIdentifiers)}" : core;
    }

    public override bool Equals(object? obj) =>
        obj is SemanticVersion other && CompareTo(other) == 0;

    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(ToString());

    public static bool operator >(SemanticVersion left, SemanticVersion right) => left.CompareTo(right) > 0;
    public static bool operator <(SemanticVersion left, SemanticVersion right) => left.CompareTo(right) < 0;
    public static bool operator >=(SemanticVersion left, SemanticVersion right) => left.CompareTo(right) >= 0;
    public static bool operator <=(SemanticVersion left, SemanticVersion right) => left.CompareTo(right) <= 0;
}

