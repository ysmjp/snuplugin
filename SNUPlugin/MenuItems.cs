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
        [MenuItem("Dobrain/Generate Prefab From XLSX")]
        private static void menuSNUPlugin()
        {
            UnityEngine.Object prevInst = GameObject.Find("SNUPlugin");
            var obj = new GameObject("SNUPlugin");
            SNUPlugin inst = obj.AddComponent<SNUPlugin>();
            inst.destroyObject(prevInst);
            inst.openDialog(); //open xlsx
            try
            {
                UnityEngine.Object.Destroy(obj); //delete itself
                UnityEngine.Object.DestroyImmediate(obj);
                UnityEngine.Object.DestroyObject(obj);
            } catch
            { }
        }
    }
}
