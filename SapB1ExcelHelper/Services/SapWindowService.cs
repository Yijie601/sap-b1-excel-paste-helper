using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SapB1ExcelHelper.Services;

public sealed record SapWindowInfo(
    nint RootWindow,
    nint ApInvoiceWindow,
    int Left,
    int Top,
    int Width,
    int Height,
    uint ProcessId,
    uint ThreadId);

public sealed class SapWindowService
{
    private const string MdiChildClass = "TMMDIChildClass";

    public bool IsSapForeground() => TryGetForegroundSapRoot(out _, out _, out _);

    public bool TryGetActiveApInvoice(out SapWindowInfo? info, out string error)
    {
        info = null;
        if (!TryGetForegroundSapRoot(out var root, out var processId, out var threadId))
        {
            error = "SAP Business One is not active. Switch to AP Invoice and press F8 again.";
            return false;
        }

        var apInvoice = FindFocusedMdiChild(threadId);
        if (apInvoice == nint.Zero)
        {
            apInvoice = FindNamedApInvoice(root);
        }

        if (apInvoice == nint.Zero || !NativeMethods.GetWindowRect(apInvoice, out var rectangle))
        {
            error = "AP Invoice window was not found. Open the AP Invoice and try again.";
            return false;
        }

        info = new SapWindowInfo(
            root,
            apInvoice,
            rectangle.Left,
            rectangle.Top,
            rectangle.Width,
            rectangle.Height,
            processId,
            threadId);
        error = string.Empty;
        return true;
    }

    public bool TryGetApInvoiceAtPoint(Point screenPoint, out SapWindowInfo? info, out string error)
    {
        info = null;
        var clickedWindow = NativeMethods.WindowFromPoint(screenPoint);
        if (clickedWindow == nint.Zero)
        {
            error = "Windows could not identify the clicked position. Try capturing the field again.";
            return false;
        }

        var root = NativeMethods.GetAncestor(clickedWindow, NativeMethods.GaRoot);
        var threadId = NativeMethods.GetWindowThreadProcessId(root, out var processId);
        if (!IsSapProcess(processId, root))
        {
            error = "The click was not inside SAP Business One. Click Capture, then click the real Supplier field in SAP.";
            return false;
        }

        var apInvoice = FindContainingApInvoice(clickedWindow);
        if (apInvoice == nint.Zero)
        {
            apInvoice = FindNamedApInvoice(root);
        }

        if (apInvoice == nint.Zero || !NativeMethods.GetWindowRect(apInvoice, out var rectangle))
        {
            error = "The click was inside SAP, but not inside an open A/P Invoice window.";
            return false;
        }

        info = new SapWindowInfo(
            root,
            apInvoice,
            rectangle.Left,
            rectangle.Top,
            rectangle.Width,
            rectangle.Height,
            processId,
            threadId);
        error = string.Empty;
        return true;
    }

    public bool TryFindOpenApInvoice(out SapWindowInfo? info)
    {
        info = null;
        var sapRoots = new List<(nint Window, uint ProcessId, uint ThreadId)>();
        NativeMethods.EnumWindows((window, _) =>
        {
            if (!NativeMethods.IsWindowVisible(window))
            {
                return true;
            }

            var threadId = NativeMethods.GetWindowThreadProcessId(window, out var processId);
            if (IsSapProcess(processId, window))
            {
                sapRoots.Add((window, processId, threadId));
            }

            return true;
        }, nint.Zero);

        foreach (var candidate in sapRoots)
        {
            var apInvoice = FindFocusedMdiChild(candidate.ThreadId);
            if (apInvoice == nint.Zero)
            {
                apInvoice = FindNamedApInvoice(candidate.Window);
            }
            if (apInvoice != nint.Zero && NativeMethods.GetWindowRect(apInvoice, out var rectangle))
            {
                info = new SapWindowInfo(
                    candidate.Window,
                    apInvoice,
                    rectangle.Left,
                    rectangle.Top,
                    rectangle.Width,
                    rectangle.Height,
                    candidate.ProcessId,
                    candidate.ThreadId);
                return true;
            }
        }

        return false;
    }

    public bool IsSameActiveInvoice(SapWindowInfo expected)
    {
        if (!TryGetActiveApInvoice(out var current, out _))
        {
            return false;
        }

        return current!.ProcessId == expected.ProcessId &&
               current.ApInvoiceWindow == expected.ApInvoiceWindow;
    }

    public string GetFocusedControlClass(SapWindowInfo window)
    {
        var info = new NativeMethods.GuiThreadInfo
        {
            Size = Marshal.SizeOf<NativeMethods.GuiThreadInfo>()
        };

        if (!NativeMethods.GetGUIThreadInfo(window.ThreadId, ref info) || info.Focus == nint.Zero)
        {
            return string.Empty;
        }

        return NativeMethods.GetClassName(info.Focus);
    }

    public bool IsResponsive(SapWindowInfo window)
    {
        var result = NativeMethods.SendMessageTimeout(
            window.RootWindow,
            NativeMethods.WmNull,
            nint.Zero,
            nint.Zero,
            NativeMethods.SmtoAbortIfHung,
            50,
            out _);
        return result != nint.Zero;
    }

    public bool IsWaitCursorVisible()
    {
        var cursorInfo = new NativeMethods.CursorInfo
        {
            Size = Marshal.SizeOf<NativeMethods.CursorInfo>()
        };

        if (!NativeMethods.GetCursorInfo(ref cursorInfo) ||
            (cursorInfo.Flags & NativeMethods.CursorShowing) == 0)
        {
            return false;
        }

        var waitCursor = NativeMethods.LoadCursor(nint.Zero, (nint)NativeMethods.IdcWait);
        var appStartingCursor = NativeMethods.LoadCursor(nint.Zero, (nint)NativeMethods.IdcAppStarting);
        return cursorInfo.Cursor == waitCursor || cursorInfo.Cursor == appStartingCursor;
    }

    private static bool TryGetForegroundSapRoot(out nint root, out uint processId, out uint threadId)
    {
        root = NativeMethods.GetForegroundWindow();
        if (root == nint.Zero)
        {
            processId = 0;
            threadId = 0;
            return false;
        }

        root = NativeMethods.GetAncestor(root, NativeMethods.GaRoot);
        threadId = NativeMethods.GetWindowThreadProcessId(root, out processId);
        return IsSapProcess(processId, root);
    }

    private static bool IsSapProcess(uint processId, nint rootWindow)
    {
        try
        {
            using var process = Process.GetProcessById((int)processId);
            var normalizedName = process.ProcessName
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace(".", string.Empty, StringComparison.Ordinal);
            if (normalizedName.Equals("SAPBusinessOne", StringComparison.OrdinalIgnoreCase) ||
                normalizedName.StartsWith("SAPBusinessOne", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var title = NativeMethods.GetWindowText(rootWindow);
            return title.Contains("SAP Business One", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static nint FindFocusedMdiChild(uint threadId)
    {
        var info = new NativeMethods.GuiThreadInfo
        {
            Size = Marshal.SizeOf<NativeMethods.GuiThreadInfo>()
        };
        if (!NativeMethods.GetGUIThreadInfo(threadId, ref info))
        {
            return nint.Zero;
        }

        var current = info.Focus != nint.Zero ? info.Focus : info.Active;
        while (current != nint.Zero)
        {
            if (NativeMethods.GetClassName(current).Equals(MdiChildClass, StringComparison.Ordinal))
            {
                var title = NativeMethods.GetWindowText(current);
                return IsApInvoiceTitle(title)
                    ? current
                    : nint.Zero;
            }

            current = NativeMethods.GetParent(current);
        }

        return nint.Zero;
    }

    private static nint FindContainingApInvoice(nint window)
    {
        var current = window;
        while (current != nint.Zero)
        {
            if (NativeMethods.GetClassName(current).Equals(MdiChildClass, StringComparison.Ordinal) &&
                IsApInvoiceTitle(NativeMethods.GetWindowText(current)))
            {
                return current;
            }

            current = NativeMethods.GetParent(current);
        }

        return nint.Zero;
    }

    private static nint FindNamedApInvoice(nint root)
    {
        nint found = nint.Zero;
        NativeMethods.EnumChildWindows(root, (window, _) =>
        {
            if (!NativeMethods.GetClassName(window).Equals(MdiChildClass, StringComparison.Ordinal))
            {
                return true;
            }

            var title = NativeMethods.GetWindowText(window);
            if (!IsApInvoiceTitle(title))
            {
                return true;
            }

            found = window;
            return false;
        }, nint.Zero);
        return found;
    }

    public static bool IsApInvoiceTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            return false;
        }

        var normalized = title
            .Replace("/", string.Empty, StringComparison.Ordinal)
            .Replace("\\", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal);
        return normalized.Contains("APInvoice", StringComparison.OrdinalIgnoreCase);
    }
}
