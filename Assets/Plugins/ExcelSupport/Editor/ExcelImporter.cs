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
using UnityEditor.AssetImporters;
using UnityEngine;

namespace ExcelSupport.Editor{
    [ScriptedImporter(0, new[] {
        "xlam", "xlsm", "xlsx", "xltm", "xltx", "xls"
    })]
    public class ExcelImporter : ScriptedImporter{
        public GoExcel goExcel;
        public override void OnImportAsset(AssetImportContext ctx){
            // var goExcel = ObjectFactory.CreateInstance<GoExcel>();
            var fileName = Path.GetFileName(ctx.assetPath);
            if (fileName.StartsWith("~") && File.Exists(Path.Combine(Path.GetFullPath(ctx.assetPath ?? ""), fileName.Replace("~", "").Replace("$","")))){
                ctx.mainObject.hideFlags = HideFlags.HideAndDontSave;
                ctx.SetMainObject(null);
                return;
            }
            goExcel = GoExcel.ReadExcel(ctx.assetPath);
            ctx.AddObjectToAsset("main", goExcel);
            ctx.SetMainObject(goExcel);
            Debug.Log("Import Success: " + goExcel);
        }
    }
}