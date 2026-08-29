namespace SapB1ExcelHelper.Services;

[Flags]
public enum HotkeyModifiers : uint
{
    None = 0,
    Alt = 0x0001,
    Control = 0x0002,
    Shift = 0x0004
}

public sealed record HotkeyDefinition(uint VirtualKey, HotkeyModifiers Modifiers)
{
    private const HotkeyModifiers SupportedModifiers =
        HotkeyModifiers.Control | HotkeyModifiers.Alt | HotkeyModifiers.Shift;

    public static HotkeyDefinition Default => new((uint)Keys.F8, HotkeyModifiers.None);

    public string DisplayText
    {
        get
        {
            var parts = new List<string>();
            if (Modifiers.HasFlag(HotkeyModifiers.Control))
            {
                parts.Add("Ctrl");
            }

            if (Modifiers.HasFlag(HotkeyModifiers.Alt))
            {
                parts.Add("Alt");
            }

            if (Modifiers.HasFlag(HotkeyModifiers.Shift))
            {
                parts.Add("Shift");
            }

            parts.Add(GetKeyName((Keys)VirtualKey));
            return string.Join(" + ", parts);
        }
    }

    public bool IsSupported(out string error)
    {
        var key = (Keys)VirtualKey;
        if (VirtualKey == 0 || VirtualKey > 0xFF || IsModifierKey(key))
        {
            error = "Press a non-modifier key.";
            return false;
        }

        if ((Modifiers & ~SupportedModifiers) != 0)
        {
            error = "This modifier combination is not supported.";
            return false;
        }

        var isFunctionKey = key is >= Keys.F1 and <= Keys.F24;
        if (!isFunctionKey &&
            !Modifiers.HasFlag(HotkeyModifiers.Control) &&
            !Modifiers.HasFlag(HotkeyModifiers.Alt))
        {
            error = "Use F1–F24, or combine another key with Ctrl or Alt.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    public static bool TryCreate(
        Keys keyCode,
        Keys modifierKeys,
        out HotkeyDefinition? definition,
        out string error)
    {
        var modifiers = HotkeyModifiers.None;
        if (modifierKeys.HasFlag(Keys.Control))
        {
            modifiers |= HotkeyModifiers.Control;
        }

        if (modifierKeys.HasFlag(Keys.Alt))
        {
            modifiers |= HotkeyModifiers.Alt;
        }

        if (modifierKeys.HasFlag(Keys.Shift))
        {
            modifiers |= HotkeyModifiers.Shift;
        }

        definition = new HotkeyDefinition((uint)(keyCode & Keys.KeyCode), modifiers);
        if (definition.IsSupported(out error))
        {
            return true;
        }

        definition = null;
        return false;
    }

    private static bool IsModifierKey(Keys key) => key is
        Keys.ControlKey or Keys.LControlKey or Keys.RControlKey or
        Keys.Menu or Keys.LMenu or Keys.RMenu or
        Keys.ShiftKey or Keys.LShiftKey or Keys.RShiftKey or
        Keys.LWin or Keys.RWin;

    private static string GetKeyName(Keys key) => key switch
    {
        >= Keys.D0 and <= Keys.D9 => ((int)key - (int)Keys.D0).ToString(),
        >= Keys.NumPad0 and <= Keys.NumPad9 => $"Num {(int)key - (int)Keys.NumPad0}",
        Keys.Oemcomma => ",",
        Keys.OemPeriod => ".",
        Keys.Oemplus => "+",
        Keys.OemMinus => "-",
        Keys.Space => "Space",
        Keys.Return => "Enter",
        Keys.Escape => "Esc",
        Keys.PageUp => "Page Up",
        Keys.PageDown => "Page Down",
        _ => key.ToString()
    };
}
