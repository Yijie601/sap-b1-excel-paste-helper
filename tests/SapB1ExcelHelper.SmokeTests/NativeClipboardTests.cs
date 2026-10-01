using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using SapB1ExcelHelper.Services;

internal static class NativeClipboardTests
{
    internal const string IntegrationFlag = "--native-clipboard-integration";
    internal const string ChildFlag = "--native-clipboard-child";
    internal const string NotRequestedMessage =
        "SKIP native Win32 clipboard integration: opt in with --native-clipboard-integration; a newly created private window station is required, with no fallback to an existing or user station.";
    private const uint CreateOnly = 1;
    // Read attributes + clipboard + create desktop + global atoms only.
    private const uint WindowStationAccess = 0x002E;
    // Read objects + create window + write objects; no switch/input rights.
    private const uint DesktopAccess = 0x0083;
    private const int UserObjectName = 2;
    private const int UserObjectFlags = 1;
    private const uint WindowStationVisible = 1;

    // A clipboard belongs to a window station, not merely a desktop:
    // https://learn.microsoft.com/en-us/windows/win32/winstation/about-window-stations-and-desktops
    // This process launcher never reads or writes the parent's user clipboard.
    internal static void RunInIsolatedProcess()
    {
        if (Environment.GetCommandLineArgs().Contains(ChildFlag, StringComparer.Ordinal))
        {
            throw new InvalidOperationException("Register native clipboard child dispatch before the normal test collection.");
        }
        var executable = Environment.ProcessPath ?? throw new InvalidOperationException("Cannot locate the smoke-test executable.");
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(executable), "dotnet", StringComparison.OrdinalIgnoreCase))
        {
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        }
        start.ArgumentList.Add(ChildFlag);

        using var process = Process.Start(start) ?? throw new InvalidOperationException("Cannot start the isolated clipboard test.");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(25_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new InvalidOperationException("The isolated clipboard test timed out; its child process was terminated.");
        }

        var diagnostic = output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Isolated clipboard test failed (exit {process.ExitCode}): {diagnostic.Trim()}");
        }
        if (!diagnostic.Contains("NATIVE_CLIPBOARD_ISOLATED_PASS", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Child dispatch was not registered; native clipboard test did not run.");
        }
        Console.WriteLine(diagnostic.Trim());
    }

    // Dispatch this before constructing/running the normal test collection.
    // If isolation cannot be verified, no clipboard operation is attempted.
    internal static int RunIsolatedChild()
    {
        if (!Environment.GetCommandLineArgs().Contains(ChildFlag, StringComparer.Ordinal))
        {
            Console.Error.WriteLine("Native clipboard tests may run only in an explicitly dispatched child process.");
            return 2;
        }

        var originalStation = Native.GetProcessWindowStation();
        var originalDesktop = Native.GetThreadDesktop(Native.GetCurrentThreadId());
        nint privateStation = 0;
        nint privateDesktop = 0;
        nint window = 0;
        var stationName = $"SapClipboardTest-{Guid.NewGuid():N}";
        var desktopName = $"ClipboardDesktop-{Guid.NewGuid():N}";
        var stationChanged = false;
        try
        {
            Require(originalStation != 0 && originalDesktop != 0, "Cannot identify inherited station/desktop; isolation refused.");
            var originalName = GetObjectName(originalStation);
            // Prefer a unique new station. If administrator naming rights are
            // unavailable, a null name lets Windows choose one for this logon
            // session. Both attempts use CWF_CREATE_ONLY; an existing station
            // is never opened or reused, even if it is noninteractive.
            // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-createwindowstationw
            privateStation = Native.CreateWindowStationW(stationName, CreateOnly, WindowStationAccess, 0);
            if (privateStation == 0)
            {
                var namedError = Marshal.GetLastWin32Error();
                privateStation = Native.CreateWindowStationW(null, CreateOnly, WindowStationAccess, 0);
                if (privateStation == 0)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(),
                        $"Private station isolation unavailable (named creation error {namedError}, automatic creation error {Marshal.GetLastWin32Error()}); no clipboard was read or written.");
                }
            }
            stationName = GetObjectName(privateStation);
            Require(stationName != originalName && !string.Equals(stationName, "WinSta0", StringComparison.OrdinalIgnoreCase),
                "Windows selected an inherited or interactive station; isolation refused.");
            RequireNative(Native.SetProcessWindowStation(privateStation), "Cannot select the private window station.");
            stationChanged = true;
            Require(GetObjectName(Native.GetProcessWindowStation()) != originalName, "Window station did not change; isolation refused.");

            privateDesktop = Native.CreateDesktopW(desktopName, null, 0, 0, DesktopAccess, 0);
            RequireNative(privateDesktop != 0, "Cannot create the private test desktop.");
            RequireNative(Native.SetThreadDesktop(privateDesktop), "Cannot select the private test desktop.");
            VerifyIsolation(stationName, desktopName);

            // No WS_VISIBLE, activation, keyboard/mouse simulation, or desktop switch.
            window = Native.CreateWindowExW(0, "STATIC", "Clipboard test owner", 0x80000000,
                0, 0, 1, 1, 0, 0, 0, 0);
            RequireNative(window != 0, "Cannot create a hidden clipboard-owner HWND.");
            VerifyIsolation(stationName, desktopName);

            var workersAttached = 0;
            // The runtime may have created default-desktop pool threads before
            // this child changed its process station. Verify/attach every
            // actual native-access thread before any clipboard call. Leave
            // the hook installed until this disposable child exits, including
            // on failure, so a delayed read cannot escape the isolation guard.
            ClipboardService.BeforeNativeAccess = () =>
            {
                Require(GetObjectName(Native.GetProcessWindowStation()) == stationName,
                    "A native clipboard worker is outside the private process station; access refused.");
                if (GetObjectName(Native.GetThreadDesktop(Native.GetCurrentThreadId())) != desktopName)
                {
                    RequireNative(Native.SetThreadDesktop(privateDesktop),
                        "Cannot attach the native clipboard worker to the private desktop; access refused.");
                    Interlocked.Increment(ref workersAttached);
                }
                VerifyIsolation(stationName, desktopName);
            };

            const string sample = "供应商 Café 🥟\t03-10-2026\t单据-01\tITEM-1\tO-HW\t15\t\t7.5\tTX7\t\t\t\tS-HW\r\n" +
                                  "供应商 Café 🥟\t03-10-2026\t单据-01\tITEM-2\tO-HW\t3\t\t0\tTX7\t\t\t\tS-HW\r\n";
            var firstSequence = WaitWithOwnerMessagePump(ClipboardService.SetTextAsync(sample, window,
                expectedSequence: ClipboardService.Sequence));
            var roundtrip = WaitWithOwnerMessagePump(ClipboardService.ReadTextAsync());
            var currentSequence = ClipboardService.Sequence;
            Require(roundtrip.Text == sample,
                $"Native Unicode TSV text mismatch: expected UTF-16 length {sample.Length}, actual length {roundtrip.Text?.Length.ToString() ?? "null"}; write sequence {firstSequence}, read sequence {roundtrip.Sequence}, current sequence {currentSequence}; private workers attached {workersAttached}. No clipboard contents logged.");
            Require(roundtrip.Sequence == firstSequence,
                $"Native clipboard sequence mismatch after exact text roundtrip: write sequence {firstSequence}, read sequence {roundtrip.Sequence}, current sequence {currentSequence}; private workers attached {workersAttached}.");

            VerifyIsolation(stationName, desktopName);
            const string replacement = "替换内容 — exact Unicode";
            var replacementSequence = WaitWithOwnerMessagePump(ClipboardService.SetTextAsync(replacement, window,
                expectedSequence: firstSequence));
            Require(replacementSequence != firstSequence, "A new native clipboard write did not change its sequence.");
            var rejected = false;
            try
            {
                WaitWithOwnerMessagePump(ClipboardService.SetTextAsync("must not overwrite", window,
                    expectedSequence: firstSequence));
            }
            catch (ClipboardChangedException)
            {
                rejected = true;
            }
            Require(rejected, "Conditional native write accepted a stale clipboard sequence.");
            var afterRejection = WaitWithOwnerMessagePump(ClipboardService.ReadTextAsync());
            Require(afterRejection.Text == replacement && afterRejection.Sequence == replacementSequence,
                "Rejected conditional write mutated the private clipboard.");

            TestContention(window, privateDesktop, stationName, desktopName, replacementSequence);
            TestSameOwnerPayloadChange(window, stationName, desktopName, "EXPECTED 🥟 altered");
            TestSameOwnerPayloadChange(window, stationName, desktopName, "REPLACED 🥟");
            VerifyIsolation(stationName, desktopName);
            Console.WriteLine("NATIVE_CLIPBOARD_ISOLATED_PASS Unicode TSV, conditional sequence, contention retry, and same-owner payload rejection; user clipboard untouched.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
        finally
        {
            if (window != 0)
            {
                _ = Native.DestroyWindow(window);
            }
            // Never restore the inherited station in this disposable child.
            // Even a timed-out delayed-render worker must remain confined to
            // the private clipboard until process termination. Windows closes
            // assigned station/desktop handles when the child exits.
            if (!stationChanged && privateStation != 0)
            {
                _ = Native.CloseWindowStation(privateStation);
            }
        }
    }

    private static void TestContention(nint window, nint desktop, string stationName, string desktopName, uint expectedSequence)
    {
        using var opened = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var busy = new ManualResetEventSlim();
        Exception? blockerError = null;
        var blocker = new Thread(() =>
        {
            var locked = false;
            try
            {
                RequireNative(Native.SetThreadDesktop(desktop), "Cannot isolate the clipboard-contention thread.");
                VerifyIsolation(stationName, desktopName);
                RequireNative(Native.OpenClipboard(0), "Cannot hold the private clipboard for contention testing.");
                locked = true;
                opened.Set();
                Require(release.Wait(TimeSpan.FromSeconds(10)), "Clipboard-contention release signal timed out.");
            }
            catch (Exception exception)
            {
                blockerError = exception;
                opened.Set();
            }
            finally
            {
                if (locked)
                {
                    _ = Native.CloseClipboard();
                }
            }
        }) { IsBackground = true };
        blocker.Start();
        Task<uint>? write = null;
        try
        {
            Require(opened.Wait(TimeSpan.FromSeconds(5)), "Contention thread did not open the private clipboard.");
            if (blockerError is not null)
            {
                throw new InvalidOperationException("Could not establish private clipboard contention.", blockerError);
            }
            VerifyIsolation(stationName, desktopName);
            var busyNotifications = 0;
            write = ClipboardService.SetTextAsync("忙碌恢复 🥟", window,
                onBusy: () =>
                {
                    Interlocked.Increment(ref busyNotifications);
                    busy.Set();
                }, expectedSequence: expectedSequence);
            Require(busy.Wait(TimeSpan.FromSeconds(5)), "Native contention retry did not report a busy clipboard.");
            Require(!write.IsCompleted && busyNotifications == 1,
                "Native contention did not enter the asynchronous retry path.");
            release.Set();
            var sequence = WaitWithOwnerMessagePump(write);
            Require(busyNotifications == 1, "Clipboard busy notification was repeated.");
            var result = WaitWithOwnerMessagePump(ClipboardService.ReadTextAsync());
            Require(result.Text == "忙碌恢复 🥟" && result.Sequence == sequence,
                "Native clipboard retry did not recover with the exact payload.");
        }
        finally
        {
            release.Set();
            // Finish the private retry and lock holder before returning.
            blocker.Join();
            if (write is not null)
            {
                try { _ = WaitWithOwnerMessagePump(write); }
                catch { }
            }
        }
        if (blockerError is not null)
        {
            throw new InvalidOperationException("Private clipboard-contention thread failed.", blockerError);
        }
    }

    private static void TestSameOwnerPayloadChange(nint window, string stationName, string desktopName, string mutated)
    {
        VerifyIsolation(stationName, desktopName);
        var isolationGuard = ClipboardService.BeforeNativeAccess ??
            throw new InvalidOperationException("The private native-access guard is required for the same-owner test.");
        var accesses = 0;
        var mutations = 0;
        uint mutatedSequence = 0;
        const string expected = "EXPECTED 🥟";
        // Test both an extended expected prefix and a same-length replacement,
        // requiring exact payload content and its terminating NUL.
        ClipboardService.BeforeNativeAccess = () =>
        {
            isolationGuard();
            if (Interlocked.Increment(ref accesses) == 2)
            {
                // With expectedSequence omitted, the first native access is
                // the outer write; this second access occurs after its close,
                // just before the ownership checkpoint opens the clipboard.
                // Recursive accesses still verify isolation but never mutate
                // again. The owning main thread pumps its private HWND while
                // this worker waits for the one competing native write.
                mutatedSequence = ClipboardService.SetTextAsync(mutated, window)
                    .GetAwaiter().GetResult();
                Interlocked.Increment(ref mutations);
            }
        };

        var rejected = false;
        try
        {
            try
            {
                WaitWithOwnerMessagePump(ClipboardService.SetTextAsync(expected, window));
            }
            catch (ClipboardChangedException)
            {
                rejected = true;
            }
        }
        finally
        {
            // Preserve the original private-desktop guard for all subsequent
            // native operations and any outstanding child-process worker.
            ClipboardService.BeforeNativeAccess = isolationGuard;
        }

        Require(mutations == 1, "Same-owner payload test did not inject exactly one competing private write.");
        Require(rejected, "The ownership checkpoint accepted a changed payload with the same owner HWND.");
        VerifyIsolation(stationName, desktopName);
        var remaining = WaitWithOwnerMessagePump(ClipboardService.ReadTextAsync());
        Require(remaining.Text == mutated && remaining.Sequence == mutatedSequence,
            "Rejecting a same-owner payload change overwrote the competing private copy or changed its sequence.");
    }

    private static void VerifyIsolation(string stationName, string desktopName)
    {
        var station = Native.GetProcessWindowStation();
        Require(station != 0 && GetObjectName(station) == stationName,
            "Private process window station verification failed; clipboard access refused.");
        var flags = new byte[12];
        RequireNative(Native.GetUserObjectInformationW(station, UserObjectFlags, flags, (uint)flags.Length, out _),
            "Cannot verify that the private station is noninteractive.");
        Require((BitConverter.ToUInt32(flags, 8) & WindowStationVisible) == 0,
            "Test station is interactive; clipboard access refused.");
        Require(GetObjectName(Native.GetThreadDesktop(Native.GetCurrentThreadId())) == desktopName,
            "Private thread desktop verification failed; clipboard access refused.");
    }

    private static T WaitWithOwnerMessagePump<T>(Task<T> task)
    {
        // Worker-thread EmptyClipboard can send WM_DESTROYCLIPBOARD to our
        // hidden owner HWND. Keep its private-desktop message thread pumping
        // while waiting, just as the real WinForms app does while awaiting.
        while (!task.IsCompleted)
        {
            while (Native.PeekMessageW(out var message, 0, 0, 0, 1))
            {
                _ = Native.DispatchMessageW(in message);
            }
            Thread.Sleep(10);
        }
        return task.GetAwaiter().GetResult();
    }

    private static string GetObjectName(nint handle)
    {
        Require(handle != 0, "Cannot inspect a null station/desktop handle.");
        var buffer = new byte[1024];
        RequireNative(Native.GetUserObjectInformationW(handle, UserObjectName, buffer, (uint)buffer.Length, out var required),
            "Cannot verify station/desktop identity.");
        return Encoding.Unicode.GetString(buffer, 0, checked((int)required)).TrimEnd('\0');
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    private static void RequireNative(bool condition, string message)
    {
        if (!condition) throw new Win32Exception(Marshal.GetLastWin32Error(), message);
    }

    private static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct Message
        {
            internal nint Window;
            internal uint Id;
            internal nuint WParam;
            internal nint LParam;
            internal uint Time;
            internal int X;
            internal int Y;
            internal uint Private;
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PeekMessageW(out Message message, nint window, uint first, uint last, uint remove);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern nint DispatchMessageW(in Message message);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern nint CreateWindowStationW(string? name, uint flags, uint access, nint security);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetProcessWindowStation(nint station);
        [DllImport("user32.dll")]
        internal static extern nint GetProcessWindowStation();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern nint CreateDesktopW(string name, string? device, nint mode, uint flags, uint access, nint security);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetThreadDesktop(nint desktop);
        [DllImport("user32.dll")]
        internal static extern nint GetThreadDesktop(uint threadId);
        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetUserObjectInformationW(nint handle, int index, [Out] byte[] information, uint length, out uint required);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern nint CreateWindowExW(uint extendedStyle, string className, string title, uint style,
            int x, int y, int width, int height, nint parent, nint menu, nint instance, nint parameter);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DestroyWindow(nint window);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool OpenClipboard(nint owner);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseClipboard();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseWindowStation(nint station);
    }
}
