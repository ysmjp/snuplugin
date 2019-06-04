using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

namespace SNUPlugin
{
    public class SNUPlugin : MonoBehaviour
    {
        private List<Proposal> myProposal = new List<Proposal>();

        private void OnStart()
        {
            Debug.Log("SNUPlugin initialized.");
        }

        //destroy unity object
        public void destroyObject(UnityEngine.Object obj)
        {
            if (obj == null) return;
            if (Application.isEditor)
                DestroyImmediate(obj);
            else
                Destroy(obj);
        }

        //open dialog and import excel proposal
        public bool openDialog()
        {
            string path = EditorUtility.OpenFilePanel("기획서 파일을 선택해주세요.", "", "xlsx");
            Parser myParser = new Parser();
            Generator myGenerator = new Generator();
            if (path.Length != 0)
            {
                if (!myParser.importSheet(myProposal, path))
                    return false;
                return myGenerator.generate(this, myProposal);
            }
            return false;
        }

    }
}
