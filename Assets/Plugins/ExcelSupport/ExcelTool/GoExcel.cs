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
using System.IO;
using System.Linq;
using System.Text;
using UnityEngine;

namespace GalForUnity.ExcelTool{
    [Serializable]
    public class GoExcel : ScriptableObject{
        [SerializeField] public string excelName;

        [SerializeField] public string excelPath;

        [SerializeField] public List<Sheet> sheets;

        [NonSerialized] private Dictionary<string, Sheet> _sheetMap;

        [NonSerialized] internal Action UndoCallBack;


        public Dictionary<string, Sheet> Sheets{
            get{
                if (_sheetMap == null || _sheetMap.Count != sheets.Count) _sheetMap = sheets.ToDictionary(x => x.sheetName);
                return _sheetMap;
            }
        }

        public string FileFullPath => Path.Combine(excelPath, excelName); 
        public Sheet this[string sheetName]{
            get => Sheets[sheetName];
            set => Sheets[sheetName] = value;
        }

        public Sheet this[int sheetIndex]{
            get => sheets[sheetIndex];
            set => sheets[sheetIndex] = value;
        }

        private void OnValidate(){ UndoCallBack?.Invoke(); }

        public static GoExcel ReadExcel(string datasetPath){
            var csStr = GoExcelNativeMethod.ToCsStr(
                GoExcelNativeMethod.GetExcel(GoExcelNativeMethod.ToCStr(datasetPath)));
            var goExcel = CreateInstance<GoExcel>();
            JsonUtility.FromJsonOverwrite(csStr, goExcel);
            return goExcel;
        }
        public static void ReadExcel(string datasetPath,GoExcel obj){
            var csStr = GoExcelNativeMethod.ToCsStr(
                GoExcelNativeMethod.GetExcel(GoExcelNativeMethod.ToCStr(datasetPath)));
            JsonUtility.FromJsonOverwrite(csStr,obj);
        }
        public static WriteState WriteExcel(GoExcel goExcel){
            var result = GoExcelNativeMethod.SetExcel(GoExcelNativeMethod.ToCStr(JsonUtility.ToJson(goExcel)));

            var writeState = (WriteState) result;
            if (writeState == WriteState.Failed) Debug.LogError("save failed");
            return writeState;
        }

        public Sheet GetSheet(string sheetName){ return Sheets[sheetName]; }
        public int GetSheetIndex(Sheet sheet){ return sheets.IndexOf(sheet); }
        public int GetSheetIndex(string sheetName){ return GetAllSheetName().IndexOf(sheetName); }

        public GoExcel DeepClone(){
            var goExcel = CreateInstance<GoExcel>();
            JsonUtility.FromJsonOverwrite(JsonUtility.ToJson(this), goExcel);
            return goExcel;
        }

        public void ShallowClone(GoExcel goExcel){
            var mySheet = sheets;
            goExcel.sheets=new List<Sheet>();
            foreach (var sheet in mySheet){
                goExcel.sheets.Add(new Sheet(sheet));
            }
            goExcel.excelName = excelName;
            goExcel._sheetMap = _sheetMap;
            goExcel.excelPath = excelPath;
        }

        public List<string> GetAllSheetName(){ return sheets.Select(x => x.sheetName).ToList(); }
        

        [Serializable]
        public class Sheet{
            [SerializeField] public string sheetName;

            [SerializeField] public List<Line> rows;

            public Sheet(string sheetName){ this.sheetName = sheetName; }

            public Sheet(Sheet sheet){
                sheetName = sheet.sheetName;
                rows=new List<Line>();
                foreach (var sheetRow in sheet.rows){
                    var line = new Line {
                        index = sheetRow.index, cells = new List<Cell>()
                    };
                    foreach (var sheetRowCell in sheetRow.cells){
                        line.cells.Add(sheetRowCell);
                    }
                    rows.Add(line);
                }
            }

            public Line this[int row]{
                get => rows[row];
                set => rows[row] = value;
            }

            public Line GetLine(int row, int col){ return rows[row]; }
        }

        [Serializable]
        public class Line{
            [SerializeField] public int index;

            [SerializeField] public List<Cell> cells;

            public Cell this[int col]{
                get => cells[col];
                set => cells[col] = value;
            }

            public Cell GetCell(int col){ return cells[col]; }

            public void Combine(Line line){
                cells.AddRange(line.cells);
            }
        }

        [Serializable]
        public class Cell{
            [SerializeField] public string text;

            [SerializeField] public int index;

            public string AsString => text;
            public int AsInt => int.TryParse(text, out var value) ? value : 0;
            public float AsFloat => float.TryParse(text, out var value) ? value : 0.0f;
            public double AsDouble => double.TryParse(text, out var value) ? value : 0.0d;
            public byte[] AsByteArray => string.IsNullOrWhiteSpace(text) ? new byte[] { } : Encoding.Default.GetBytes(text);

            public override string ToString(){ return AsString; }
        }
    }

    public enum WriteState{
        Success,
        Failed
    }
}