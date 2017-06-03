//
// LockAndSleep.cs
//
// Author:
//       M.A. (enmoku) <>
//
// Copyright (c) 2017 M.A. (enmoku)
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
using System.IO;

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
					Application.ExitThread();
					break;
			}
		}

		protected override void WndProc(ref Message m)
		{
			if (m.Msg == WM_SYSCOMMAND)
			{
				if (m.WParam.ToInt32() == SC_MONITORPOWER)
				{
					Console.WriteLine("Monitor power state: {0}", ((PowerMode)m.LParam.ToInt32()).ToString());

					MonitorPower?.Invoke(this, new MonitorPowerEventArgs(m.LParam.ToInt32() == 0 ? PowerMode.Off : PowerMode.On));
				}
			}
			else if (m.Msg == WM_POWERBROADCAST)
			{
				if (m.WParam.ToInt32() == PBT_POWERSETTINGCHANGE)
				{
					POWERBROADCAST_SETTING ps = (POWERBROADCAST_SETTING)Marshal.PtrToStructure(m.LParam, typeof(POWERBROADCAST_SETTING));
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
			}

			base.WndProc(ref m); // is this necessary?
		}

		public static void SleepDisplays(PowerMode powermode)
		{
			int NewPowerMode = (int)powermode; // -1 = Powering On, 1 = Low Power (low backlight, etc.), 2 = Power Off
			IntPtr Handle = new IntPtr(-1); // -1 = 0xFFFF = HWND_BROADCAST
			SendMessage(Handle, PowerManager.WM_SYSCOMMAND, PowerManager.SC_MONITORPOWER, NewPowerMode);		}

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

		[DllImport("user32.dll")]
		static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
	}

	public class LockAndSleep
	{
		public static void PrintHeader()
		{
			Console.WriteLine("Lock & Sleep - To lock workstation and put monitors to sleep.");
			Console.WriteLine("https://github.com/enmoku/lockandsleep");
			Console.WriteLine();
		}

		public static int MinMax(int value, int min, int max)
		{
			return Math.Max(Math.Min(value, max), min);
		}

		public static int delay = 2500; // 2.5 seconds
		public static bool quiet = false;
		public static bool retry = false;
		public static int retrywait = 60000 * 5; // 5 minutes
		public const int retrywaitmax = 60000 * 30; // 30 minutes
		public const int retrywaitmin = 60000 * 1; // 1 minute

		public static int powermode = 2; // 2 = off, 1 = standby, -1 = on

		public static bool locked = false;

		public static PowerManager pman;

		static bool CurrentLockState = false;
		static PowerMode CurrentPowerState = PowerMode.Invalid;

		static Timer RetrySleepMode;

		public static void Main()
		{
			string[] args = Environment.GetCommandLineArgs();
			for (var i = 1; i < args.Length; i++)
			{
				switch (args[i])
				{
					case "-w":
					case "--wait":
						if (i + 1 < args.Length)
							delay = Convert.ToInt32(args[++i]);
						break;
					case "-q":
					case "--quiet":
						quiet = true;
						break;
					case "-r":
					case "--retry":
						retry = true;
						if (i + 1 < args.Length)
						{
							if (!args[i + 1].StartsWith("-", StringComparison.InvariantCulture))
								retrywait = MinMax(Convert.ToInt32(args[++i]), retrywaitmin, retrywaitmax);
						}
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
						Console.WriteLine("--retry [ms]   Retry sleep with # ms delay [default: 5 minutes]");
						Console.WriteLine("-r [ms]");
						// Standby puts on lower power mode, like dimming brightness
						//Console.WriteLine("--standby     Set monitor to standby instead of off.");
						//Console.WriteLine("-s");
						Console.WriteLine("--help        Show help and exit with no action.");
						Console.WriteLine("-h");
						return;
				}
			}

			// DEBUG
			//retry = true;
			//retrywait = 5000;
			//powermode = 1;

			if (!quiet)
			{
				PrintHeader();

				Console.WriteLine("Delay: {0}ms", delay);

				if (retry) Console.WriteLine("Retry: {0}ms", retrywait);
			}

			if (retry)
			{
				pman = new PowerManager();
				// TODO: Move as much as possible of the following to PowerManagementHwnd
				RetrySleepMode = new Timer();
				RetrySleepMode.Interval = retrywait;
				int retrycount = 0;
				int reinforcecount = 0;
				RetrySleepMode.Tick += (sender, e) =>
				{
					// Stop trying after a while
					if (retrycount > 5 && (retrycount/2) > reinforcecount)
					{
						RetrySleepMode.Stop();
						// There's no point keeping this app running.
						Application.ExitThread();
						return;
					}

					if (CurrentPowerState == ((PowerMode)powermode))
					{
						RetrySleepMode.Stop();
						return;
					}

					retrycount += 1;

					LASTINPUTINFO info = new LASTINPUTINFO();
					info.cbSize = (uint)Marshal.SizeOf(info);
					info.dwTime = 0;
					bool n = GetLastInputInfo(ref info);
					if (n)
					{
						float eticks = Convert.ToSingle(Environment.TickCount);
						float uticks = Convert.ToSingle(info.dwTime);
						float idletime = (eticks - uticks) / 1000;
						if (idletime < Convert.ToInt32(retrywait / 1000))
						{
							if (!quiet) Console.WriteLine("User active too recently ({0}s). Skipping sleep reinforcement.", idletime);
							return;
						}

						System.Diagnostics.Debug.WriteLine("Idle time: {0:N1}s", idletime);
					}
					else
					{
						System.Diagnostics.Debug.WriteLine("Failure to get idle time, cancelling sleep reinforcement.");
						return;
					}

					if (!quiet) Console.WriteLine("Powering down again.");
					PowerManager.SleepDisplays((PowerMode)powermode);
					reinforcecount += 1;
				};

				pman.MonitorPower += (sender, e) =>
				{
					CurrentPowerState = e.Mode;
					if (CurrentPowerState == PowerMode.On)
						RetrySleepMode.Start();
				};
				pman.SessionLock += (sender, e) =>
				{
					CurrentLockState = e.Locked;
					if (CurrentLockState == false)
						RetrySleepMode.Stop();
				};
			}

			if (delay > 0) System.Threading.Thread.Sleep(delay);

			if (!quiet)
			{
				Console.WriteLine();
				Console.WriteLine("Locking...");
			}
			LockWorkStation();

			if (!quiet) Console.WriteLine("Powering down...");
			PowerManager.SleepDisplays((PowerMode)powermode);

			if (retry)
			{
				if (!quiet) Console.WriteLine("Waiting for unlock...");
				Application.Run();
				if (!quiet) Console.WriteLine("Unlocked, exiting.");
			}

			if (!quiet)
			{
				Console.WriteLine();
				Console.WriteLine("Done.");
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
