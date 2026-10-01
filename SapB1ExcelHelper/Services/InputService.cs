using System.Diagnostics;

namespace SapB1ExcelHelper.Services;

internal static class InputService
{
    private const uint InputMouse = 0;
    private const uint InputKeyboard = 1;
    private const uint MouseLeftDown = 0x0002;
    private const uint MouseLeftUp = 0x0004;
    private const uint KeyUp = 0x0002;
    private const ushort VkControl = 0x11;

    internal static void Click(int x, int y)
    {
        if (!NativeMethods.SetCursorPos(x, y))
        {
            throw new InvalidOperationException("Unable to move the mouse to the calibrated SAP field.");
        }

        try
        {
            NativeMethods.EnsureInputSent(new[] { Mouse(MouseLeftDown), Mouse(MouseLeftUp) });
        }
        catch
        {
            _ = NativeMethods.SendInput(1, new[] { Mouse(MouseLeftUp) },
                System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.Input>());
            throw;
        }
    }

    internal static void SelectAll() => SendShortcut(0x41);

    internal static void Paste() => SendShortcut(0x56);

    internal static Task WaitForModifiersReleasedAsync(CancellationToken cancellationToken) =>
        WaitForModifiersReleasedAsync(ModifiersPressed, cancellationToken);

    internal static async Task WaitForModifiersReleasedAsync(Func<bool> modifiersPressed, CancellationToken cancellationToken)
    {
        var timer = Stopwatch.StartNew();
        while (modifiersPressed())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (timer.Elapsed >= TimeSpan.FromSeconds(3))
            {
                throw new SapAutomationException("Release Ctrl, Alt, Shift and Windows keys before running the paste.");
            }
            await Task.Delay(25, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static bool ModifiersPressed() => NativeMethods.GetAsyncKeyState(0x10) < 0 ||
        NativeMethods.GetAsyncKeyState(VkControl) < 0 || NativeMethods.GetAsyncKeyState(0x12) < 0 ||
        NativeMethods.GetAsyncKeyState(0x5B) < 0 || NativeMethods.GetAsyncKeyState(0x5C) < 0;

    private static void SendShortcut(ushort key)
    {
        if (ModifiersPressed())
        {
            throw new SapAutomationException("A modifier key is held down. Check SAP before starting another run.");
        }

        var inputs = new[]
        {
            Key(VkControl, false),
            Key(key, false),
            Key(key, true),
            Key(VkControl, true)
        };
        try
        {
            NativeMethods.EnsureInputSent(inputs);
        }
        catch
        {
            // A partial SendInput may leave our Ctrl/key down. Release only
            // those two keys, without replaying the shortcut or its paste.
            _ = NativeMethods.SendInput(2, new[] { Key(key, true), Key(VkControl, true) },
                System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.Input>());
            throw;
        }
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
}
