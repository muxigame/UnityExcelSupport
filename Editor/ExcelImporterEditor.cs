//======================================================================
//
//       CopyRight 2019-2022 © MUXI Game Studio 
//       . All Rights Reserved 
//
//        FileName :  ExcelImporterEditor.cs
//
//        Created by 半世癫(Roc) at 2022-05-24 02:27:31
//
//======================================================================

using System.Collections.Generic;
using System.IO;
using System.Linq;
using GalForUnity.ExcelTool;
using UnityEditor;
using UnityEditor.AssetImporters;
using UnityEngine;
using UnityEngine.UIElements;

namespace ExcelSupport.Editor{
    [CustomEditor(typeof(ExcelImporter))]
    // [CanEditMultipleObjects]
    public class ExcelImporterEditor : ScriptedImporterEditor{
        private GoExcel _assetExcel;
        private GoExcel _goExcel;
        private string _itemString;
        private VisualElement _root;
        private ScrollView _scrollView;

        private SerializedProperty _serializedProperty;
        private ImportSetting.ExcelImportSetting _unImportSettings;

        public override void OnEnable(){
            base.OnEnable();
            _assetExcel = ((ExcelImporter) target).assetExcel;
            _unImportSettings = ImportSetting.instance.GetImportSetting(_assetExcel);
            if(!_goExcel) _goExcel = GoExcel.ReadExcel(Path.Combine(_assetExcel.excelPath,_assetExcel.excelName));
            _unImportSettings.Init(_goExcel);
            _goExcel.UndoCallBack += DrawTable;
            _serializedProperty = serializedObject.FindProperty("assetExcel");
        }

        public override void OnDisable(){
            base.OnDisable();
            _goExcel.UndoCallBack -= DrawTable;
        }

        private IEnumerable<string> GetAssetPaths(){ return targets.OfType<AssetImporter>().Select(i => i.assetPath); }

        protected void DrawTable(){
            var allSheetName = _goExcel.GetAllSheetName();
            if (allSheetName == null ||allSheetName.Count ==0) return;
            var sheet = string.IsNullOrEmpty(_itemString) ? _goExcel[allSheetName[0]] : _goExcel[_itemString];
            DrawTable(sheet);
        }

        protected override void Apply(){
            _assetExcel = _unImportSettings.RemoveUnImport(_goExcel,_assetExcel);
            // GoExcel.WriteExcel(_goExcel);
            ImportSetting.Save();
            base.Apply();
        }

        public override VisualElement CreateInspectorGUI(){
            if (_scrollView == null) _scrollView = new ScrollView(ScrollViewMode.Horizontal);
            if (_root == null)  _root = new VisualElement {
                style = {
                    flexDirection = FlexDirection.Column
                }
            };
            var allSheetName = _goExcel.GetAllSheetName();
            if (allSheetName == null||allSheetName.Count==0) return _root;
            
            var sheet = string.IsNullOrEmpty(_itemString) ? _goExcel[allSheetName[0]] : _goExcel[_itemString];
            DrawTable(sheet);
            _itemString = sheet?.sheetName;

            _root.Add(_scrollView);
            _root.Add(new Button {
                text = "Apply", clickable = new Clickable(() => { ApplyAndImport(); })
            });
            _root.Add(new Button {
                text = "Revert",
                clickable = new Clickable(() => {
                    _goExcel = GoExcel.ReadExcel(Path.Combine(_assetExcel.excelPath,_assetExcel.excelName));
                    DrawTable(_goExcel[_itemString]);
                })
            });
            _root.Add(new Button {
                text = "导出num文件",
                clickable = new Clickable(()=> {
                    _assetExcel = _unImportSettings.RemoveUnImport(_goExcel,_assetExcel);
                    ImportSetting.Save();
                    _assetExcel.WriteNum();
                })
            });
            return _root;
        }

        protected void DrawTable(GoExcel.Sheet sheet){
            if (_scrollView == null) _scrollView = new ScrollView(ScrollViewMode.Horizontal);
            _scrollView.Clear();
            var sheetIndex = _goExcel.GetSheetIndex(sheet.sheetName);
            var table = new VisualElement {
                style = {
                    flexDirection = FlexDirection.Column
                }
            };
            for (var i = 0; i < sheet.rows.Count; i++){
                var rowIndex = i;
                if (i == 0){
                    var unImportLine = new VisualElement {
                        style = {
                            flexDirection = FlexDirection.Row
                        }
                    };
                    var tableToggle = new Toggle {
                        style = {
                            width = 70
                        },
                        value = _unImportSettings[sheetIndex].sheetImportInfo
                    };
                    tableToggle.RegisterValueChangedCallback(x => _unImportSettings[sheetIndex].sheetImportInfo = x.newValue);
                    unImportLine.Add(tableToggle);
                    for (var j = 0; j < _unImportSettings[sheetIndex].colImportInfo.Count; j++){
                        var cellIndex = j;
                        var rowToggle = new Toggle {
                            style = {
                                width = 70
                            },
                            value = _unImportSettings[sheetIndex].colImportInfo[cellIndex]
                        };
                        rowToggle.RegisterValueChangedCallback(x => _unImportSettings[sheetIndex].colImportInfo[cellIndex] = x.newValue);
                        unImportLine.Add(rowToggle);
                    }

                    table.Add(unImportLine);
                }

                var lines = sheet.rows;
                var line = new VisualElement {
                    style = {
                        flexDirection = FlexDirection.Row
                    }
                };
                for (var j = 0; j < lines[i].cells.Count; j++){
                    var cell = sheet[i][j];
                    if (j == 0){
                        var colToggle = new Toggle {
                            value = _unImportSettings[sheetIndex].rowImportInfo[rowIndex]
                        };
                        colToggle.RegisterValueChangedCallback(x => { _unImportSettings[sheetIndex].rowImportInfo[rowIndex] = x.newValue; });
                        line.Add(colToggle);
                    }

                    var textField = new TextField {
                        value = cell.text,
                        style = {
                            width = 70
                        }
                    };
                    textField.RegisterValueChangedCallback(x => {
                        Undo.RecordObject(_goExcel, "textField");
                        cell.text = x.newValue;
                    });
                    line.Add(textField);
                }

                table.Add(line);
            }

            _scrollView.contentContainer.Add(table);
        }

        protected override void OnHeaderGUI(){
            base.OnHeaderGUI();
            serializedObject.Update();
            if (_root != null) _root.style.maxWidth = EditorGUIUtility.currentViewWidth - 20;
            if (EditorGUILayout.DropdownButton(new GUIContent(_itemString), FocusType.Keyboard)){
                var menu = new GenericMenu();
                foreach (var item in _goExcel.GetAllSheetName()){
                    if (string.IsNullOrEmpty(item)) continue;
                    //添加菜单
                    menu.AddItem(new GUIContent(item), _itemString.Equals(item), x => {
                        var userData = (GoExcel.Sheet) x;
                        _itemString = userData.sheetName;
                        DrawTable(userData);
                    }, _goExcel[item]);
                }

                menu.ShowAsContext(); //显示菜单
            }

            EditorGUILayout.PropertyField(_serializedProperty);
            serializedObject.SetIsDifferentCacheDirty();
            ApplyRevertGUI();
        }
        
    }
}