//======================================================================
//
//       CopyRight 2019-2022 © MUXI Game Studio 
//       . All Rights Reserved 
//
//        FileName :  ExcelImporter.cs
//
//        Created by 半世癫(Roc) at 2022-05-20 00:43:14
//
//======================================================================

using System.IO;
using GalForUnity.ExcelTool;
using UnityEditor;
using UnityEditor.AssetImporters;
using UnityEngine;

namespace ExcelSupport.Editor{
    [ScriptedImporter(0, new[] {
        "xlam", "xlsm", "xlsx", "xltm", "xltx", "xls"
    })]
    public class ExcelImporter : ScriptedImporter{
        public GoExcel assetExcel;
        
        public override void OnImportAsset(AssetImportContext ctx){
            // var goExcel = ObjectFactory.CreateInstance<GoExcel>();
            var fileName = Path.GetFileName(ctx.assetPath);
            if (fileName.StartsWith("~") && File.Exists(Path.Combine(Path.GetFullPath(ctx.assetPath ?? ""), fileName.Replace("~", "").Replace("$", "")))){
                ctx.mainObject.hideFlags = HideFlags.HideAndDontSave;
                ctx.SetMainObject(null);
                return;
            }
            if (!assetExcel){
                var loadAssetAtPath = AssetDatabase.LoadAssetAtPath<GoExcel>(ctx.assetPath);
                if (loadAssetAtPath){
                    assetExcel = loadAssetAtPath;
                } else{
                    assetExcel = ScriptableObject.CreateInstance<GoExcel>();
                    GoExcel.ReadExcel(ctx.assetPath, assetExcel);
                    var importSetting = ImportSetting.instance;
                    var excelImportSetting = importSetting.GetImportSetting(AssetDatabase.GUIDFromAssetPath(ctx.assetPath));
                    assetExcel=excelImportSetting.RemoveUnImport(assetExcel);
                }
                // var excelImportSetting = ImportSetting.instance.GetImportSetting(assetExcel);
                // assetExcel=excelImportSetting.RemoveUnImport(assetExcel.DeepClone());
                // Debug.Log(excelImportSetting);
            }
            ctx.AddObjectToAsset("main", assetExcel);
            ctx.SetMainObject(assetExcel);
            Debug.Log("Import Success: " + assetExcel);
        }
    }

    public class EditorCallBack : UnityEditor.AssetModificationProcessor{
        // public static AssetDeleteResult OnWillDeleteAsset(string assetPath,RemoveAssetOptions option)
        // {
        //     // Debug.LogFormat("delete : {0}", assetPath);
        //     // var loadAssetAtPath = AssetDatabase.LoadAssetAtPath<GoExcel>(assetPath);
        //     // if (loadAssetAtPath is not null){
        //     //     ImportSetting.instance.
        //     // }
        //     // AssetDeleteResult.DidNotDelete;
        //     return AssetDeleteResult.DidDelete;
        // }
    }
}