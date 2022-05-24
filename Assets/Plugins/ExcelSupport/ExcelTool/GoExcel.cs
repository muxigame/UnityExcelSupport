//======================================================================
//
//       CopyRight 2019-2022 © MUXI Game Studio 
//       . All Rights Reserved 
//
//        FileName :  GoExcel.cs
//
//        Created by 半世癫(Roc) at 2022-05-24 00:38:25
//
//======================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UnityEngine;
using Object = UnityEngine.Object;

namespace GalForUnity.ExcelTool{
    [Serializable]
    public class GoExcel : ScriptableObject
    {
        [SerializeField]
        public string excelName;  
        [SerializeField]
        public string excelPath;
        [SerializeField]
        public List<Sheet> sheets;
        [NonSerialized]
        private Dictionary<string, Sheet> _sheetMap;
        [NonSerialized]
        public Action UndoCallBack;
        private void OnValidate(){
            UndoCallBack?.Invoke();
        }
        public Dictionary<string, Sheet> Sheets{
            get{
                if ((_sheetMap == null) || _sheetMap.Count != sheets.Count){
                    _sheetMap = sheets.ToDictionary(x => x.sheetName);
                }
                return _sheetMap;
            }
        }
        public static GoExcel ReadExcel(string datasetPath)
        {
            var csStr = GoExcelNativeMethod.ToCsStr(
                GoExcelNativeMethod.GetExcel(GoExcelNativeMethod.ToCStr(datasetPath)));
            var goExcel = ScriptableObject.CreateInstance<GoExcel>();
            JsonUtility.FromJsonOverwrite(csStr,goExcel);
            return goExcel;
        }    
        public static WriteState WriteExcel(GoExcel goExcel)
        {
            var result = GoExcelNativeMethod.SetExcel(GoExcelNativeMethod.ToCStr(JsonUtility.ToJson(goExcel)));

            WriteState writeState = (WriteState) result;
            if (writeState == WriteState.Failed){
                Debug.LogError("save failed");
            }
            return writeState;
        }    
        
        public Sheet GetSheet(string sheetName){
            return Sheets[sheetName];
        }
        public Sheet this[string sheetName]{
            get=>Sheets[sheetName];
            set=>Sheets[sheetName]=value;
        }      
        public Sheet this[int sheetIndex]{
            get=>sheets[sheetIndex];
            set=>sheets[sheetIndex]=value;
        }
        
        public List<string> GetAllSheetName() => sheets.Select(x => x.sheetName).ToList();
        
        [Serializable]
        public class Sheet{
            public Sheet(string sheetName){
                this.sheetName=sheetName;
            }
            
            [SerializeField]
            public string sheetName;
            [SerializeField]
            public List<Line> rows;

            public Line GetLine(int row, int col){
                return rows[row];
            }
            public Line this[int row]{
                get=>rows[row];
                set=>rows[row]=value;
            }
        }

        [Serializable]
        public class Line{
            public Line(){ }
            [SerializeField]
            public int index;
            [SerializeField]
            public List<Cell> cells;

            public Cell this[int col]{
                get=>cells[col];
                set=>cells[col]=value;
            }

            public Cell GetCell(int col){
                return cells[col];
            }
        }
        [Serializable]
        public class Cell{
            
            [SerializeField]
            public string text;
            [SerializeField]
            public int index;

            public string AsString => text;
            public int AsInt => int.TryParse(text, out int value) ? value : 0;
            public float AsFloat => float.TryParse(text, out float value) ? value : 0.0f;
            public double AsDouble => double.TryParse(text, out double value) ? value : 0.0d;
            public byte[] AsByteArray => string.IsNullOrWhiteSpace(text)?new byte[]{}:Encoding.Default.GetBytes(text);

            public override string ToString(){
                return AsString;
            }
        }
    }

    public enum WriteState{
        Success,
        Failed
    }
}