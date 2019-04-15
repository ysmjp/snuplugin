using System;
using UnityEngine;
using UnityEditor;

namespace SNUPlugin
{
    public class MenuItems
    {
        [MenuItem("Dobrain/SNUPlugin")]
        private static void menuSNUPlugin()
        {
            SNUPlugin myInst = new SNUPlugin();
            myInst.requestSheets();
        }
    }
}
