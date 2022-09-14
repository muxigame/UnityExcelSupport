//======================================================================
//
//       CopyRight 2019-2022 © MUXI Game Studio 
//       . All Rights Reserved 
//
//        FileName :  ExcelHandler.cs
//
//        Created by 半世癫(Roc) at 2022-05-24 22:32:53
//
//======================================================================

using GalForUnity.ExcelTool;
using UnityEngine;

namespace ExcelSupport.Demo{
    public class ExcelHandler : MonoBehaviour{
        public GoExcel goExcel;

        private void Awake(){
            for (var i = 0; i < goExcel[0].rows.Count; i++){
                for (var j = 0; j < goExcel[0].rows[i].cells.Count; j++) Debug.Log(goExcel[0][i][j].text);
            }
        }
    }
}