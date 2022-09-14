//======================================================================
//
//       CopyRight 2019-2022 © MUXI Game Studio 
//       . All Rights Reserved 
//
//        FileName : ImportSetting.cs   Time : 2022-05-26 23:14:30
//
//======================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using GalForUnity.ExcelTool;
using UnityEditor;
using UnityEngine;

namespace ExcelSupport.Editor{
    [Serializable]
    [CreateAssetMenu(fileName = "ImportSetting")]
    [FilePathAttribute("ExcelSupport/ImportSetting.meta",FilePathAttribute.Location.ProjectFolder)]
    public class ImportSetting : ScriptableSingleton<ImportSetting>{
        [SerializeField]
        private List<ImportConfig> tableImportSettings;
        private Dictionary<GUID, ExcelImportSetting> _importSettings;
        
        private Dictionary<GUID, ExcelImportSetting> ImportSettings{
            get{
                if (_importSettings == null || _importSettings.Count != tableImportSettings.Count){
                    _importSettings=new Dictionary<GUID, ExcelImportSetting>();
                    foreach (var tableImportSetting in tableImportSettings){
                        GUID.TryParse(tableImportSetting.goExcelGuid, out GUID guid);
                        _importSettings.Add(guid,tableImportSetting.excelImportSetting);
                    }
                }
                return _importSettings;
            }
        }
        public ExcelImportSetting GetImportSetting(GoExcel goExcel){
            var guidFromAssetPath = AssetDatabase.GUIDFromAssetPath(AssetDatabase.GetAssetPath(goExcel));
            return GetImportSetting(guidFromAssetPath);
        }
        public ExcelImportSetting GetImportSetting(GUID guid){
            if (tableImportSettings == null) tableImportSettings = new List<ImportConfig>();
            if (this.ImportSettings.ContainsKey(guid)) return this.ImportSettings[guid];
            var excelImportSetting = new ExcelImportSetting();
            // ImportSettings.Add(guidFromAssetPath,excelImportSetting);
            this.tableImportSettings.Add(new ImportConfig {
                goExcelGuid = guid.ToString(), excelImportSetting = excelImportSetting
            });
            return this.ImportSettings[guid];
        }
        public static void Save(){ instance.Save(true); }

        [Serializable]
        public class ImportConfig{
            [SerializeField]
            public string goExcelGuid;
            [SerializeField]
            public ExcelImportSetting excelImportSetting;
        }

        [Serializable]
        public class ExcelImportSetting{
            [SerializeField]
            public List<SheetImportSetting> sheetImportSettings = new List<SheetImportSetting>();

            public SheetImportSetting this[int index] => sheetImportSettings[index];

            public GoExcel RemoveUnImport(GoExcel goExcel,GoExcel shallowClone){
                goExcel.ShallowClone(shallowClone);
                while (sheetImportSettings.Count < shallowClone.sheets.Count) sheetImportSettings.Add(new SheetImportSetting());
                Init(shallowClone);
                for (var sheetIndex = shallowClone.sheets.Count - 1; sheetIndex >= 0; sheetIndex--){
                    var unImportSetting = this[sheetIndex];
                    this[sheetIndex].sheetName = shallowClone.GetAllSheetName()[sheetIndex];

                    if (!unImportSetting.sheetImportInfo){
                        shallowClone.sheets.RemoveAt(sheetIndex);
                        continue;
                    }
                    for (var i = unImportSetting.colImportInfo.Count - 1; i >= 0; i--){ //The First bool is excelData
                        for (var j = shallowClone[sheetIndex].rows.Count - 1; j >= 0; j--)
                            if (!unImportSetting.colImportInfo[i] && shallowClone[sheetIndex].rows[j].cells.Count > i)
                                shallowClone[sheetIndex].rows[j].cells.RemoveAt(i);
                    }

                    for (var i = unImportSetting.rowImportInfo.Count - 1; i >= 0; i--)
                        if (!unImportSetting.rowImportInfo[i])
                            shallowClone[sheetIndex].rows.RemoveAt(i);
                }
                return shallowClone;
            }
            public GoExcel RemoveUnImport(GoExcel shallowClone){
                while (sheetImportSettings.Count < shallowClone.sheets.Count) sheetImportSettings.Add(new SheetImportSetting());
                Init(shallowClone);
                for (var sheetIndex = shallowClone.sheets.Count - 1; sheetIndex >= 0; sheetIndex--){
                    var unImportSetting = this[sheetIndex];
                    this[sheetIndex].sheetName = shallowClone.GetAllSheetName()[sheetIndex];
                    if (!unImportSetting.sheetImportInfo){
                        shallowClone.sheets.RemoveAt(sheetIndex);
                        continue;
                    }
                    for (var i = unImportSetting.colImportInfo.Count - 1; i >= 0; i--){ //The First bool is excelData
                        for (var j = shallowClone[sheetIndex].rows.Count - 1; j >= 0; j--)
                            if (!unImportSetting.colImportInfo[i] && shallowClone[sheetIndex].rows[j].cells.Count > i)
                                shallowClone[sheetIndex].rows[j].cells.RemoveAt(i);
                    }

                    for (var i = unImportSetting.rowImportInfo.Count - 1; i >= 0; i--)
                        if (!unImportSetting.rowImportInfo[i])
                            shallowClone[sheetIndex].rows.RemoveAt(i);
                }
                return shallowClone;
            }
            public ExcelImportSetting Init(GoExcel goExcel){
                // if (this.sheetImportSettings != null && this.sheetImportSettings.Count == goExcel.sheets.Count) return this;
                if(this.sheetImportSettings==null)this.sheetImportSettings=new List<SheetImportSetting>();
                while (this.sheetImportSettings.Count < goExcel.sheets.Count) this.sheetImportSettings.Add(new ExcelImportSetting.SheetImportSetting());
                for (var sheetIndex = 0; sheetIndex < goExcel.sheets.Count; sheetIndex++){
                    var unImportSetting = this[sheetIndex];
                    this[sheetIndex].sheetName = goExcel.GetAllSheetName()[sheetIndex];
                    unImportSetting.sheetImportInfo = true;
                    var rowsCount = goExcel[sheetIndex].rows.Count;
                    if (unImportSetting.rowImportInfo == null) unImportSetting.rowImportInfo = Enumerable.Repeat(true, rowsCount).ToList();
                    while (unImportSetting.rowImportInfo.Count <rowsCount){
                        unImportSetting.rowImportInfo.Add(true);
                    }
                    while (unImportSetting.rowImportInfo.Count >rowsCount){
                        unImportSetting.rowImportInfo.RemoveAt(unImportSetting.rowImportInfo.Count-1);
                    }
                    var rowCol = goExcel[sheetIndex].rows.Max(x => x.cells.Count);
                    if (unImportSetting.colImportInfo == null) unImportSetting.colImportInfo = Enumerable.Repeat(true, rowCol).ToList();
                    while(unImportSetting.colImportInfo.Count<rowCol) 
                        unImportSetting.colImportInfo.Add(true);
                    while (unImportSetting.colImportInfo.Count >rowCol){
                        unImportSetting.colImportInfo.RemoveAt(unImportSetting.colImportInfo.Count-1);
                    }
                }
                return this;
            }

            public override string ToString(){
                return $"colImportInfo{sheetImportSettings.Count},colImportInfo:{sheetImportSettings[0]?.colImportInfo.Count}";
            }

            [Serializable]
            public class SheetImportSetting{
                [SerializeField] public List<bool> colImportInfo;

                [SerializeField] public List<bool> rowImportInfo;

                [SerializeField] public bool sheetImportInfo = true;

                [SerializeField] public string sheetName;
            }
        }
    }
}