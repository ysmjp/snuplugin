using System;
using UnityEngine;
using UnityEditor;

namespace SNUPlugin
{
    public class HookObject
    {
        public static string ReturnString()
        {
            return "ReturnedString";
        }

        public string ReturnInstanceString()
        {
            return "InstanceString";
        }
    }
}
