using System;
using System.Runtime.InteropServices;
using System.Security;

namespace Win32API {
	public class KnownFolder {
		public readonly HResult Result;
		public readonly string Path;

		public KnownFolder(string guidText, uint flags) {
			IntPtr pszPath = IntPtr.Zero;
			try {
				this.Result = (HResult)SHGetKnownFolderPath(new Guid(guidText), flags, IntPtr.Zero, out pszPath);
				this.Path = Marshal.PtrToStringAuto(pszPath);
			} finally {
				if (pszPath != IntPtr.Zero) Marshal.FreeCoTaskMem(pszPath);
			}
		}

		[DllImport("shell32.dll"), SuppressUnmanagedCodeSecurity]
		static extern int SHGetKnownFolderPath(
			[MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);
	}

	public enum HResult {
		OK = 0,
		Fail = -2147467259, // 0x80004005
		NotFound = -2147024894, // 0x80070002
		AccessDenied = -2147024891, // 0x80070005
		InvalidArg = -2147024809, //0x80070057
	}
}
