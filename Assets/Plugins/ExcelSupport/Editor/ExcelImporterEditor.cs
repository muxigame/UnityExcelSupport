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
        private GoExcel _excel;
        private string _itemString;
        private VisualElement _root;

        private ScrollView _scrollView;

        private SerializedProperty _serializedProperty;
        // override extraDataType to return the type that will be used in the Editor.
        // protected override Type extraDataType => typeof(GoExcel);

        public override void OnEnable(){
            base.OnEnable();
            _excel = (GoExcel) assetTarget;
            _excel.UndoCallBack += DrawTable;
            _serializedProperty = serializedObject.FindProperty("goExcel");
        }

        public override void OnDisable(){
            base.OnDisable();
            // ReSharper disable once DelegateSubtraction
            _excel.UndoCallBack -= DrawTable;
        }

        private IEnumerable<string> GetAssetPaths(){ return targets.OfType<AssetImporter>().Select(i => i.assetPath); }

        protected void DrawTable(){
            var sheet = string.IsNullOrEmpty(_itemString) ? _excel[_excel.GetAllSheetName().FirstOrDefault()] : _excel[_itemString];
            DrawTable(sheet);
        }

        protected override void Apply(){
            base.Apply();
            GoExcel.WriteExcel(((ExcelImporter) target).goExcel);
        }

        public override VisualElement CreateInspectorGUI(){
            _scrollView = new ScrollView(ScrollViewMode.Horizontal);
            var sheet = string.IsNullOrEmpty(_itemString) ? _excel[_excel.GetAllSheetName().FirstOrDefault()] : _excel[_itemString];
            DrawTable(sheet);
            _itemString = sheet?.sheetName;
            _root = new VisualElement {
                style = {
                    flexDirection = FlexDirection.Column
                }
            };
            _root.Add(_scrollView);
            _root.Add(new Button {
                text = "Apply", clickable = new Clickable(() => { ApplyAndImport(); })
            });
            _root.Add(new Button {
                text = "Revert",
                clickable = new Clickable(() => {
                    ImportAssets(GetAssetPaths());
                    DrawTable(_excel[_itemString]);
                })
            });
            return _root;
        }

        protected void DrawTable(GoExcel.Sheet sheet){
            _scrollView.Clear();
            var table = new VisualElement {
                style = {
                    flexDirection = FlexDirection.Column
                }
            };
            for (var i = 0; i < sheet.rows.Count; i++){
                var lines = sheet.rows;
                var line = new VisualElement {
                    style = {
                        flexDirection = FlexDirection.Row
                    }
                };
                for (var j = 0; j < lines[i].cells.Count; j++){
                    var cell = sheet[i][j];
                    var textField = new TextField {
                        value = cell.text,
                        style = {
                            width = 70
                        }
                    };
                    textField.RegisterValueChangedCallback(x => {
                        Undo.RecordObject(_excel, "textField");
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
            _root.style.maxWidth = EditorGUIUtility.currentViewWidth - 20;
            if (EditorGUILayout.DropdownButton(new GUIContent(_itemString), FocusType.Keyboard)){
                var menu = new GenericMenu();
                foreach (var item in _excel.GetAllSheetName()){
                    if (string.IsNullOrEmpty(item)) continue;
                    //添加菜单
                    menu.AddItem(new GUIContent(item), _itemString.Equals(item), x => {
                        var userData = (GoExcel.Sheet) x;
                        _itemString = userData.sheetName;
                        DrawTable(userData);
                    }, _excel[item]);
                }

                menu.ShowAsContext(); //显示菜单
            }

            EditorGUILayout.PropertyField(_serializedProperty);
            serializedObject.SetIsDifferentCacheDirty();
            ApplyRevertGUI();
        }

        private static void ImportAssets(IEnumerable<string> paths){
            foreach (var path in paths) AssetDatabase.WriteImportSettingsIfDirty(path);
            AssetDatabase.StartAssetEditing();
            foreach (var path in paths) AssetDatabase.ImportAsset(path);
            AssetDatabase.StopAssetEditing();
        }
    }
}