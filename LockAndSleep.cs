//
// LockAndSleep.cs
//
// Author:
//       M.A. (https://github.com/mkahvi)
//
// Copyright (c) 2017–2019 M.A.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;

namespace LockAndSleepWorkstation
{
	public enum PowerMode
	{
		On = -1,
		Off = 2,
		Standby = 1,
		Invalid = 0
	}

	public class MonitorPowerEventArgs : EventArgs
	{
		public PowerMode Mode;
		public MonitorPowerEventArgs(PowerMode mode) { Mode = mode; }
	}

	public class SessionLockEventArgs : EventArgs
	{
		public bool Locked = false;
		public SessionLockEventArgs(bool locked = false) { Locked = locked; }
	}

	public class PowerManager : Form
	{
		public const int WM_SYSCOMMAND = 0x0112;
		public const int WM_POWERBROADCAST = 0x218;
		public const int SC_MONITORPOWER = 0xF170;
		public const int PBT_POWERSETTINGCHANGE = 0x8013;
		public const int HWND_BROADCAST = 0xFFFF;

		public event EventHandler<SessionLockEventArgs> SessionLock;
		public event EventHandler<MonitorPowerEventArgs> MonitorPower;

		public PowerManager()
		{
			RegisterPowerSettingNotification(Handle, ref GUID_CONSOLE_DISPLAY_STATE, DEVICE_NOTIFY_WINDOW_HANDLE);
			SystemEvents.SessionSwitch += SessionSwitchHandler;
		}

		public void SessionSwitchHandler(object sender, SessionSwitchEventArgs e)
		{
			switch (e.Reason)
			{
				case SessionSwitchReason.SessionLock:
					// Expected
					System.Diagnostics.Debug.WriteLine("Session locked.");
					SessionLock?.Invoke(this, new SessionLockEventArgs(true));
					break;
				case SessionSwitchReason.SessionUnlock:
					System.Diagnostics.Debug.WriteLine("Session unlocked.");
					SessionLock?.Invoke(this, new SessionLockEventArgs(false));
					goto default;
				default:
					// Anything else, we bail out to avoid interfering with anything.
					Console.WriteLine("Unrecognized session state, bailing out to avoid interference.");
					Console.WriteLine("Event: " + e.Reason.ToString());
					// this can include sessions logon/logofff and many others
					Application.ExitThread();
					break;
			}
		}

		protected override void WndProc(ref Message m)
		{
			if (m.Msg == WM_POWERBROADCAST && m.WParam.ToInt32() == PBT_POWERSETTINGCHANGE)
			{
				var ps = (POWERBROADCAST_SETTING)Marshal.PtrToStructure(m.LParam, typeof(POWERBROADCAST_SETTING));
				if (ps.PowerSetting == GUID_CONSOLE_DISPLAY_STATE)
				{
					switch (ps.Data)
					{
						case 0x0:
							MonitorPower?.Invoke(this, new MonitorPowerEventArgs(PowerMode.Off));
							break;
						case 0x1:
							MonitorPower?.Invoke(this, new MonitorPowerEventArgs(PowerMode.On));
							break;
							/*
						case 0x2:
							pm = PowerMode.Standby;
							MonitorPower?.Invoke(this, new MonitorPowerEventArgs(PowerMode.Standby));
							break;
							*/
					}
				}
			}

			base.WndProc(ref m); // is this necessary?
		}

		public static void SetMode(PowerMode powermode)
		{
			int NewPowerMode = (int)powermode; // -1 = Powering On, 1 = Low Power (low backlight, etc.), 2 = Power Off
			var Handle = new IntPtr(HWND_BROADCAST);
			var result = new IntPtr(-1); // unused, but necessary
			uint timeout = 200; // ms per window, we don't really care if they process them
			SendMessageTimeoutFlags flags = SendMessageTimeoutFlags.SMTO_ABORTIFHUNG;
			SendMessageTimeout(Handle, WM_SYSCOMMAND, SC_MONITORPOWER, NewPowerMode, flags, timeout, out result);
		}

		//static Guid GUID_MONITOR_POWER_ON = new Guid(0x02731015, 0x4510, 0x4526, 0x99, 0xE6, 0xE5, 0xA1, 0x7E, 0xBD, 0x1A, 0xEA);
		// GUID_CONSOLE_DISPLAY_STATE is supposedly Win8 and newer, but works on Win7 too for some reason.
		static Guid GUID_CONSOLE_DISPLAY_STATE = new Guid(0x6fe69556, 0x704a, 0x47a0, 0x8f, 0x24, 0xc2, 0x8d, 0x93, 0x6f, 0xda, 0x47);

		// http://www.pinvoke.net/default.aspx/user32.registerpowersettingnotification
		const int DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000;
		[DllImport("user32.dll", SetLastError = true, EntryPoint = "RegisterPowerSettingNotification", CallingConvention = CallingConvention.StdCall)]
		static extern IntPtr RegisterPowerSettingNotification(IntPtr hRecipient, ref Guid PowerSettingGuid, Int32 Flags);
		[StructLayout(LayoutKind.Sequential, Pack = 4)]
		internal struct POWERBROADCAST_SETTING
		{
			public Guid PowerSetting;
			public uint DataLength;
			public byte Data;
		}

		[Flags]
		public enum SendMessageTimeoutFlags : uint
		{
			SMTO_NORMAL = 0x0,
			SMTO_BLOCK = 0x1,
			SMTO_ABORTIFHUNG = 0x2,
			SMTO_NOTIMEOUTIFNOTHUNG = 0x8,
			SMTO_ERRORONEXIT = 0x20
		}

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern IntPtr SendMessageTimeout(
			IntPtr hWnd,
			uint Msg,
			int wParam,
			int lParam,
			SendMessageTimeoutFlags flags,
			uint timeout,
			out IntPtr result);
	}

	public class LockAndSleep
	{
		public static void PrintHeader()
		{
			Console.WriteLine("Lock & Sleep - To lock workstation and put monitors to sleep.");
			Console.WriteLine("https://github.com/mkahvi/lockandsleep");
			Console.WriteLine();
		}

		public static int MinMax(int value, int min, int max) => Math.Max(Math.Min(value, max), min);

		public static int Delay = 2500; // 2.5 seconds
		public static bool Quiet = false;
		public static bool Retry = false;
		public static bool Pause = false;
		public static bool WaitForUnlock = false;
		public static int RetryWait = 60000 * 5; // 5 minutes
		public const int RetryWaitMax = 60000 * 30; // 30 minutes
		public const int RetryWaitMin = 60000 * 1; // 1 minute

		public static PowerMode PowerModeIntent = PowerMode.Off;

		public static bool WorkstationLockState = false;

		public static PowerManager pman;

		static bool CurrentLockState = false;
		static PowerMode CurrentPowerState = PowerMode.Invalid;

		static Timer RetrySleepMode;

		public static void WriteLine(string message) => Console.WriteLine("[{0}] {1}", DateTime.Now.TimeOfDay, message);

		static bool Awaken { get; set; } = false;

		public static void Main()
		{
			string[] args = Environment.GetCommandLineArgs();
			for (var i = 1; i < args.Length; i++)
			{
				switch (args[i])
				{
					case "--awaken":
					case "-a":
						Awaken = true;
						break;
					case "-w":
					case "--wait":
						if (i + 1 < args.Length)
							Delay = Convert.ToInt32(args[++i]);
						break;
					case "-q":
					case "--quiet":
						Quiet = true;
						break;
					case "-r":
					case "--retry":
						Retry = true;
						if (i + 1 < args.Length)
						{
							if (!args[i + 1].StartsWith("-", StringComparison.InvariantCulture))
								RetryWait = MinMax(Convert.ToInt32(args[++i]), RetryWaitMin, RetryWaitMax);
						}
						break;
					case "--pause":
					case "-p":
						Pause = true;
						break;
					case "--unlock":
					case "-u":
						WaitForUnlock = true;
						break;
					/*
					// Redundant cases
					case "-h":
					case "--help":
					*/
					default:
						PrintHeader();
						Console.WriteLine("--wait ms     Wait # ms before locking and sleeping.");
						Console.WriteLine("-w ms");
						// Not Implemented
						Console.WriteLine("--retry [ms]  Retry sleep with # ms delay [default: 5 minutes]");
						Console.WriteLine("-r [ms]");
						Console.WriteLine("--unlock      Wait indefinitely for workstation unlock.");
						Console.WriteLine("-u");
						Console.WriteLine("--pause       Pause at end of execution");
						Console.WriteLine("-p");
						// Standby puts on lower power mode, like dimming brightness
						//Console.WriteLine("--standby     Set monitor to standby instead of off.");
						//Console.WriteLine("-s");
						Console.WriteLine("--help        Show help and exit with no action.");
						Console.WriteLine("-h");
						return;
				}
			}

			if (Awaken)
			{
				WriteLine("Waking monitors");
				PowerManager.SetMode(PowerMode.On);
				System.Threading.Thread.Sleep(500); // BUG: for lols
				return;
			}

			// DEBUG
			//retry = true;
			//retrywait = 5000;
			//PowerModeIntent = PowerMode.Standby; // For debugging

			if (!Quiet)
			{
				PrintHeader();

				// Output active settings.
				Console.WriteLine(string.Format("Delay: {0}ms", Delay));
				if (Retry) Console.WriteLine(string.Format("Retry: {0}ms", RetryWait));
				Console.WriteLine("Wait for unlock: " + (WaitForUnlock ? "Enabled" : "Disabled"));
			}

			if (Retry)
			{
				pman = new PowerManager();
				// TODO: Move as much as possible of the following to PowerManagementHwnd
				RetrySleepMode = new Timer();
				RetrySleepMode.Interval = RetryWait;
				int retrycount = 0;
				int reinforcecount = 0;
				RetrySleepMode.Tick += (sender, e) =>
				{
					// Stop trying after a while
					if (!WaitForUnlock)
					{
						if (retrycount > 5 && (retrycount / 2) > reinforcecount)
						{
							RetrySleepMode.Stop();
							// There's no point keeping this app running.
							Application.ExitThread();
							return;
						}
					}

					if (CurrentPowerState == PowerModeIntent)
					{
						RetrySleepMode.Stop();
						return;
					}

					retrycount += 1;

					var info = new LASTINPUTINFO();
					info.cbSize = (uint)Marshal.SizeOf(info);
					info.dwTime = 0;
					bool n = GetLastInputInfo(ref info);
					if (n)
					{
						float eticks = Convert.ToSingle(Environment.TickCount);
						float uticks = Convert.ToSingle(info.dwTime);
						float idletime = (eticks - uticks) / 1000;
						if (idletime < Convert.ToInt32(RetryWait / 1000))
						{
							if (!Quiet) WriteLine(string.Format("User active too recently ({0:N1}s ago).", idletime));
							return;
						}

						System.Diagnostics.Debug.WriteLine("Idle time: {0:N1}s", idletime);
					}
					else
					{
						System.Diagnostics.Debug.WriteLine("Failure to get idle time, cancelling sleep reinforcement.");
						return;
					}

					if (!Quiet) WriteLine("Powering down again.");
					PowerManager.SetMode(PowerModeIntent);
					reinforcecount += 1;
				};

				pman.MonitorPower += (sender, e) =>
				{
					PowerMode OldPowerState = CurrentPowerState;
					CurrentPowerState = e.Mode;
					if (!Quiet) WriteLine(string.Format("Monitor power state: {0}", CurrentPowerState));
					if (CurrentPowerState == PowerMode.On)
						RetrySleepMode.Start();

					if (!Retry && OldPowerState != PowerMode.On && CurrentPowerState == PowerMode.On)
					{
						Application.ExitThread();
					}
				};
				pman.SessionLock += (sender, e) =>
				{
					CurrentLockState = e.Locked;
					if (!Quiet) WriteLine(string.Format("Session locked: {0}", CurrentLockState));
					if (CurrentLockState == false)
						RetrySleepMode.Stop();
				};
			}

			if (Delay > 0) System.Threading.Thread.Sleep(Delay);

			if (!Quiet)
			{
				Console.WriteLine();
				WriteLine("Locking...");
			}
			LockWorkStation();

			if (!Quiet) WriteLine("Powering down...");
			PowerManager.SetMode(PowerModeIntent);

			if (Retry)
			{
				if (!Quiet) WriteLine("Waiting for unlock...");
				Application.Run();
				if (!Quiet) WriteLine("Unlocked, exiting.");
			}

			if (!Quiet)
			{
				Console.WriteLine();
				WriteLine("Done.");
			}

			if (Pause && !Quiet)
			{
				WriteLine("Press any key to end.");
				Console.ReadKey();
			}
		}

		[DllImport("user32.dll")]
		public static extern bool LockWorkStation();

		[DllImport("user32.dll")]
		static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

		[StructLayout(LayoutKind.Sequential)]
		struct LASTINPUTINFO
		{
			public static readonly int SizeOf = Marshal.SizeOf(typeof(LASTINPUTINFO));

			[MarshalAs(UnmanagedType.U4)]
			public UInt32 cbSize;
			[MarshalAs(UnmanagedType.U4)]
			public UInt32 dwTime;
		}
	}
}
