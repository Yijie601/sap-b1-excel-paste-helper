namespace SapB1ExcelHelper.Services;

internal static class InputService
{
    private const uint InputMouse = 0;
    private const uint InputKeyboard = 1;
    private const uint MouseLeftDown = 0x0002;
    private const uint MouseLeftUp = 0x0004;
    private const uint KeyUp = 0x0002;
    private const uint KeyUnicode = 0x0004;
    private const ushort VkControl = 0x11;
    private const ushort VkTab = 0x09;

    internal static void Click(int x, int y)
    {
        if (!NativeMethods.SetCursorPos(x, y))
        {
            throw new InvalidOperationException("Unable to move the mouse to the calibrated SAP field.");
        }

        NativeMethods.EnsureInputSent(new[]
        {
            Mouse(MouseLeftDown),
            Mouse(MouseLeftUp)
        });
    }

    internal static void SelectAll() => SendShortcut(0x41);

    internal static void Paste() => SendShortcut(0x56);

    internal static void Tab() => NativeMethods.EnsureInputSent(new[]
    {
        Key(VkTab, false),
        Key(VkTab, true)
    });

    internal static void SendUnicodeText(string value)
    {
        if (value.Length == 0)
        {
            return;
        }

        var inputs = new List<NativeMethods.Input>(value.Length * 2);
        foreach (var character in value)
        {
            inputs.Add(UnicodeKey(character, false));
            inputs.Add(UnicodeKey(character, true));
        }

        NativeMethods.EnsureInputSent(inputs.ToArray());
    }

    private static void SendShortcut(ushort key)
    {
        NativeMethods.EnsureInputSent(new[]
        {
            Key(VkControl, false),
            Key(key, false),
            Key(key, true),
            Key(VkControl, true)
        });
    }

    private static NativeMethods.Input Mouse(uint flags) => new()
    {
        Type = InputMouse,
        Data = new NativeMethods.InputUnion
        {
            Mouse = new NativeMethods.MouseInput { Flags = flags }
        }
    };

    private static NativeMethods.Input Key(ushort virtualKey, bool keyUp) => new()
    {
        Type = InputKeyboard,
        Data = new NativeMethods.InputUnion
        {
            Keyboard = new NativeMethods.KeyboardInput
            {
                VirtualKey = virtualKey,
                Flags = keyUp ? KeyUp : 0
            }
        }
    };

    private static NativeMethods.Input UnicodeKey(char character, bool keyUp) => new()
    {
        Type = InputKeyboard,
        Data = new NativeMethods.InputUnion
        {
            Keyboard = new NativeMethods.KeyboardInput
            {
                ScanCode = character,
                Flags = KeyUnicode | (keyUp ? KeyUp : 0)
            }
        }
    };
}

