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
            UnityEngine.Object prevInst = GameObject.Find("SNUPlugin");
            var obj = new GameObject("SNUPlugin");
            SNUPlugin inst = obj.AddComponent<SNUPlugin>();
            inst.destroyObject(prevInst);
            inst.openDialog(); //open xlsx
            inst.destroyObject(inst); //delete itself
        }
    }
}
