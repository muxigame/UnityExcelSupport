using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace GalForUnity.ExcelTool{
    public static class GoExcelNativeMethod{
        private const string DLLName = "GoExcel.dll";

        [DllImport(DLLName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr GetSheetNames(byte[] excelPath);

        [DllImport(DLLName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr GetExcel(byte[] excelPath);
        
        [DllImport(DLLName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int SetExcel(byte[] excelPath);

        [DllImport(DLLName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr GetSheet(byte[] excelPath, byte[] sheetName);

        public static byte[] ToCStr(string str){
            var bytes = new List<byte>();
            bytes.AddRange(Encoding.UTF8.GetBytes(str));
            bytes.AddRange(BitConverter.GetBytes('\0'));
            return bytes.ToArray();
        }

        public static string ToCsStr(IntPtr ptr){
            var bytes = new List<byte>();
            var i = 0;
            var b = Marshal.ReadByte(ptr, i);
            while (b != '\0'){
                bytes.Add(b);
                i++;
                b = Marshal.ReadByte(ptr, i);
            }

            return Encoding.UTF8.GetString(bytes.ToArray());
        }
    }
}