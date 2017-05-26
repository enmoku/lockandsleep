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

namespace LockAndSleepWorkstation
{
	public class LockAndSleep
	{
		public static void PrintHeader()
		{
			Console.WriteLine("Lock & Sleep - To lock workstation and put monitors to sleep.");
			Console.WriteLine("https://github.com/enmoku/lockandsleep");
			Console.WriteLine();
		}

		public static void Main()
		{
			int delay = 2500;
			int retry = 15000;
			bool quiet = false;
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
					case "-r":
					case "--retry":
						if (i + 1 < args.Length)
							retry = Convert.ToInt32(args[++i]);
						break;
					case "-q":
					case "--quiet":
						quiet = true;
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
						//Console.WriteLine("--retry ms    Retry locking and sleeping after # ms.");
						//Console.WriteLine("-r ms");
						Console.WriteLine("--help        Show help and exit with no action.");
						Console.WriteLine("-h");
						return;
				}
			}

			if (!quiet)
			{
				PrintHeader();

				Console.WriteLine("Delay: {0}ms", delay);
				//Console.WriteLine("Retry: {0}ms", retry);
			}

			if (delay > 0)
				System.Threading.Thread.Sleep(delay);

			if (!quiet)
			{
				Console.WriteLine();
				Console.WriteLine("Locking...");
			}
			LockWorkStation();

			uint Message = 0x0112; // 0x0112 = WM_SYSCOMMAND
			int Parameter1 = 0xF170; // 0xF170 = SC_MONITORPOWER
			int Parameter2 = 2; // -1 = Powering On, 1 = Low Power (low backlight, etc.), 2 = Power Off
			IntPtr Handle = new IntPtr(-1); // -1 = 0xFFFF = HWND_BROADCAST
			if (!quiet) Console.WriteLine("Sleeping...");
			SendMessage(Handle, Message, Parameter1, Parameter2);

			if (!quiet)
			{
				Console.WriteLine();
				Console.WriteLine("Done.");
			}
		}

		[DllImport("user32.dll")]
		public static extern bool LockWorkStation();

		[DllImport("user32.dll")]
		static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
	}
}
