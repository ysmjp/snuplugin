using System;
using UnityEngine;
using UnityEditor;
using System.Collections;
using UnityEngine.Networking;
using System.Runtime.InteropServices;

namespace SNUPlugin
{
    public class MenuItems
    {
        [MenuItem("Dobrain/SNUPlugin")]
        private static void menuSNUPlugin()
        {
            var obj = new GameObject("SNUPlugin");
            SNUPlugin inst = obj.AddComponent<SNUPlugin>();
            inst.openDialog();
            //var obj = new GameObject("SNUPlugin");
            //SNUPlugin inst = obj.AddComponent<SNUPlugin>();
            //var inst = new SNUPlugin();
            //inst.requestSheets();
            //MonoBehaviour.DestroyImmediate(obj);
            //EditorUtility.DisplayDialog("", Basic(), "");
        }

        //[DllImport(".\\Assets\\Plugins\\ExternalGoogleAPI.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        //private static extern string Basic();


    }
}
