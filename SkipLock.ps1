cls

#============================================================================
# Type Declarations
#============================================================================

Add-Type -TypeDefinition @"
	using System;
	using System.Runtime.InteropServices;

	public class Display {
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetDpiForWindow(IntPtr hWnd);
	}
"@;

Add-Type -TypeDefinition @"
	using System;
	using System.Diagnostics;
	using System.Runtime.InteropServices;

	public class GlobalEventHooks {
		public delegate IntPtr HookProc (int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr GetModuleHandle(string lpModuleName);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr GetForegroundWindow();

		// Hook IDs
		public const int WH_KEYBOARD_LL = 13;
		public const int WH_MOUSE_LL = 14;

		// Hook Codes
		public static IntPtr _keyboardHookID = IntPtr.Zero;
		public static IntPtr _mouseHookID = IntPtr.Zero;

		public static HookProc _keyboardProc = HookCallback;
		public static HookProc _mouseProc = HookCallback;

		public static Action UpdateUserActivity;

		public static void SetKeyboardHook () {
			_keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0);
		}

		public static void SetMouseHook () {
			_mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0);
		}

		// Callback
		private static IntPtr HookCallback (int nCode, IntPtr wParam, IntPtr lParam) {
			if (nCode >= 0) {
				// UpdateUserActivity ().Invoke ();
                UpdateUserActivity ();
			}
			return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
		}

		// Unhook
		public static void UnhookKeyboard () {
			UnhookWindowsHookEx(_keyboardHookID);
			UnhookWindowsHookEx(_mouseHookID);
		}
	}

	public class ComputerMouse
	{
		[DllImport("user32.dll")]
		public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

		[DllImport("user32.dll")]
		public static extern void SetCursorPos(int x, int y);

		public const int MOUSEEVENTF_LEFTDOWN = 0x02;
		public const int MOUSEEVENTF_LEFTUP = 0x04;

		public static void SetCursorTo(int x, int y)
		{
			SetCursorPos(x, y);
		}

		public static void Click(int x, int y)
		{
			mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, x, y, 0, 0);
		}
	}
"@;

Add-Type -TypeDefinition @"
	using System;
	using System.Drawing;
	using System.Text;
	using System.Runtime.InteropServices;

	public class User32 {
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr GetDC(IntPtr hWnd);

		[DllImport("Gdi32.dll")]
		public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

		public static int GetDpiScalingFactor ()
		{
			const int LOG_PIXELS_X = 88;
			IntPtr hdc = GetDC(IntPtr.Zero);
			int dpi = GetDeviceCaps(hdc, LOG_PIXELS_X);
			ReleaseDC(IntPtr.Zero, hdc);
			return dpi;
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool SetForegroundWindow(IntPtr hWnd);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

		[StructLayout(LayoutKind.Sequential)]
		public struct WINDOWPLACEMENT {
			public int length;
			public int flags;
			public int showCmd;
			public POINT ptMinPosition;
			public POINT ptMaxPosition;
			public RECT rcNormalPosition;
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct POINT {
			public int x;
			public int y;
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct RECT {
			public int left;
			public int top;
			public int right;
			public int bottom;
		}
	}
"@;

$global:activityDetected = $false;

[GlobalEventHooks]::UpdateUserActivity = {
	$global:activityDetected = $true;
	$global:lastUserActivityDateTime = Get-Date;
}

[GlobalEventHooks]::SetKeyboardHook();
[GlobalEventHooks]::SetMouseHook();

$lastUserActivityDateTime = Get-Date;

#============================================================================
# Getting OS-set DPI scaling (typically, > 100%) on
# the actual display resolution for the same of better
# readability and accuracy in the mouse cursor positioning.
#============================================================================

$hwnd = (Get-Process -Id $pid).MainWindowHandle;
$dpi = [Display]::GetDpiForWindow($hwnd);
$osSetDPIScalingFactor = $dpi / 96 * 100;

#============================================================================
# Getting the actual and scaled display resolution
#============================================================================

# Find the largest monitor, and get its real screen size.

[double]$realWidth = 0;
[double]$realHeight = 0;

$videoControllers = Get-WmiObject -Class "Win32_VideoController";

foreach ($videoController in $videoControllers) {
	$width = $videoController.CurrentHorizontalResolution;
	$height = $videoController.CurrentVerticalResolution;

	Write-Host "Identified Monitor $($videoController.Name) with Resolution ($width, $height)";

	if ($width -gt $realWidth) {
		$realWidth = $width;
		$realHeight = $height;
	}
}

# Calculate the scaled resolution.
$screen = [System.Windows.Forms.Screen]::PrimaryScreen;
[double]$scaledWidth = $screen.Bounds.Width;
[double]$scaledHeight = $screen.Bounds.Height;

$realResolution = "($realWidth, $realHeight)";
$scaledResolution = "($scaledWidth, $scaledHeight)";

Write-Host "=============================================================";
Write-Host "Real Resolution: $realResolution";
Write-Host "Scaled Resolution: $scaledResolution";
Write-Host "";

[double] $widthFactor = $scaledWidth / $realWidth;
[double] $heightFactor = $scaledHeight / $realHeight;

Write-Host "Width Factor: $widthFactor";
Write-Host "Height Factor: $heightFactor";
Write-Host "";

#============================================================================
# Compute zoom percentage
#============================================================================

Write-Host "OS-set DPI Scaling Factor: $osSetDPIScalingFactor %";
Write-Host "Display Zoom percentage: $(1 / $heightFactor * 100) %";
Write-Host ""

switch ($realResolution) {
	"(3840, 2160)" {
		Write-Host "The display screen is Home 4K Monitor";
	}
	"(1920, 2160)" {
		Write-Host "The display screen is Home 4K Monitor in Split Mode";
	}
	"(1920, 1080)" {
		Write-Host "The display screen is Office or Built-in HD Monitor";
	}
}

Write-Host "=============================================================";
Write-Host "";
Write-Host "";

[int] $x = 0;
[int] $y = 0;

$x = ($scaledWidth * 0.85) * ($osSetDPIScalingFactor / 100);
$y = ($scaledHeight - 30) * ($osSetDPIScalingFactor / 100);

Write-Host "Click Position is ($x, $y)";

While ($true) {
	$currentTime = Get-Date -Format "hh:mm tt";
	Write-Host "";

	$counter ++;
	Write-Host "Iteration # $counter, current time is $currentTime.";

	#=============================================================================
	# Identifying current cursor position
	#=============================================================================

	$currentX = [System.Windows.Forms.Cursor]::Position.X;
	$currentY = [System.Windows.Forms.Cursor]::Position.Y;

	#=============================================================================
	# Moving the cursor to the target position
	#=============================================================================

	$secondsToWaitFurther = [Math]::Floor($(New-TimeSpan -Start ([DateTime]::Now) -End $global:lastUserActivityDateTime.AddMinutes(2)).TotalSeconds);

	if ($secondsToWaitFurther -lt 0) {
		$secondsToWaitFurther = 0;
	}

	if ($global:activityDetected -eq $false -or $secondsToWaitFurther -eq 0) {
		[ComputerMouse]::SetCursorTo($x, $y);
		[ComputerMouse]::Click($x, $y);

		Start-Sleep -Seconds 120;
	} else {
		Write-Host "Skipping, as user activity detected. Waiting for $secondsToWaitFurther seconds ...";
		$global:activityDetected = $false;
		Start-Sleep -Seconds $secondsToWaitFurther;
	}
}