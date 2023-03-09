using System;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
class NDSINI
{
    [DllImport("KERNEL32.DLL", EntryPoint = "GetPrivateProfileStringW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
    private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, string lpReturnString, int nSize, string lpFilename);
    [DllImport("KERNEL32.DLL", EntryPoint = "WritePrivateProfileStringW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
    private static extern int WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFilename);
    [DllImport("KERNEL32.DLL")]
    public static extern long GetSystemDefaultLCID();

    [DllImport("KERNEL32.DLL", EntryPoint = "SetLocaleInfoA", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
    public static extern int SetLocaleInfo(long Locale, long LCType, string lpLCData);
    public const long LOCALE_SSHORTDATE = (long)0x1F;

    public static string GetINI(string IniFile, string Category, string Key, string DefaultValue)
    {
        string ReturnString = new string(' ', 1024);
        GetPrivateProfileString(Category, Key, DefaultValue, ReturnString, 1024, IniFile);
        return ReturnString.Split('\0')[0];
    }
    public static string WriteINI(string IniFile, string Category, string Key, string Value)
    {
        int rs;
        NDSINI.ParseCrLf(ref Value);
        rs = WritePrivateProfileString(Category, Key, Value, IniFile);
        return rs.ToString();
    }
    private static void ParseCrLf(ref string Value)
    {
        Value.Replace("\n", "");
        Value.Replace("\r", "");
    }
}