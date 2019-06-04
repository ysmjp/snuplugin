using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEditor;

namespace SNUPlugin
{
    //prefab generator
    class Generator
    {
        //generate prefab
        public bool generate(SNUPlugin snuplugin, List<Proposal> proposalList)
        {
            string strDir, strPath, strObjectName, strQManagerName, strQManagerPath;
            int intSuccess = 0;
            Proposal prop;
            GameObject obj;
            if (proposalList.Count == 0) return false;
            //strDir = "Assets/Contents/Resources/Weekday/" + proposalList[0].ContentsIndex.ToString();
            strDir = "Assets/Contents/Resources/Weekday/105";
            if (!mkdir(strDir)) return false; //fail to make direcotry
            for (int i = 0; i < proposalList.Count; i++) //loop thru proposal list
            {
                prop = proposalList[i];
                strObjectName = "ch" + prop.ContentsIndex.ToString() + "_" + prop.QuestionIndex.ToString();
                strPath = strDir + "/" + strObjectName + ".prefab";
                strQManagerName = getQManagerName(prop);
                strQManagerPath = $"Assets/Contents/Scripts/Question/Weekday/{strQManagerName}.cs";
                Debug.Log($"SNUPlugin: Trying to instantiate {strQManagerPath}");
                if (strQManagerName == "undefined")
                {
                    obj = new GameObject(strObjectName);
                }
                else
                {
                    //obj = UnityEngine.Object.Instantiate(GameObject.Find(strQManager)); //clone question manager
                    //obj = Resources.Load<GameObject>(strQManagerPath);
                    //obj = new GameObject(strObjectName);
                    //obj.AddComponent(Resources.Load(strQManagerPath).GetType());

                    //obj = AssetDatabase.LoadAssetAtPath<GameObject>(strQManagerPath);

                    //obj = (GameObject)UnityEngine.Object.Instantiate(Resources.Load(strQManagerPath)); //clone question manager
                    obj = UnityEngine.Object.Instantiate(AssetDatabase.LoadAssetAtPath<UnityEngine.GameObject>(strQManagerPath)); //clone question manager
                    obj.name = strObjectName;
                    obj.GetComponent("Explanation Board").name = strQManagerName;
                }
                createPrefab(obj, strPath); //create and save prefab
                snuplugin.destroyObject(obj);
                intSuccess++;
            }
            EditorUtility.DisplayDialog("SNUPlugin", $"{Path.GetFileName(proposalList[0].Filename)}을 로드하여\n{proposalList[0].ContentsIndex.ToString()}번 컨텐츠에 해당하는 {intSuccess}개의 Prefab을 생성하였습니다.\n{strDir} 폴더를 확인해주세요.", "닫기");
            return true;
        }

        //get question manager name
        public string getQManagerName(Proposal proposal)
        {
            //BlockDropQManager
            //CardFlipQManager
            //ClickInOrderQManager
            //ClickMultipleQManager
                //ClickToStackQManager
            //DetectionQManager
            //DragAndDropQManager
            //DrawingQManager
            //LinePatternQManager
                //PolymorphQManager
           //QuestionManager
           //WhackAMoleQManager
                //AssistRacingQManager
            //ClickOneQManager
            //ConnectLinesQManager
                //DiverseClickQManager
                //DiverseDragAndDropQManager
            //ErasingQManager
                //GateBarrierQManager
                //RotationCirclePuzzleQManager
                //SamePictureQManager
                //SlidePuzzleQManager
                //StackQManager

            switch (proposal.GameType)
            {
                case DobrainGameType.Undefined:
                    return "undefined";//default
                case DobrainGameType.None:
                    return "undefined";//default
                case DobrainGameType.ChoiceOne:
                    return "ClickOneQManager";
                case DobrainGameType.MoleMoving:
                    return "WhackAMoleQManager";
                case DobrainGameType.MoleStatic:
                    return "WhackAMoleQManager";
                case DobrainGameType.MoleAnimating:
                    return "WhackAMoleQManager";
                case DobrainGameType.DrawingImage:
                    return "DrawingQManager";
                case DobrainGameType.DrawingPattern:
                    return "LinePatternQManager";
                case DobrainGameType.DragAndDrop:
                    return "DragAndDropQManager";
                case DobrainGameType.ChoiceMulti:
                    return "ClickMultipleQManager";
                case DobrainGameType.Erase:
                    return "ErasingQManager";
                case DobrainGameType.Catch:
                    return "DetectionQManager";
                case DobrainGameType.DrawingLine:
                    return "ConnectLinesQManager";
                case DobrainGameType.Click:
                    return "ClickInOrderQManager";
                case DobrainGameType.PilingBlocks:
                    return "BlockDropQManager";
                case DobrainGameType.FlipingCards:
                    return "CardFlipQManager";
                case DobrainGameType.Other:
                    return "undefined";//default
                default:
                    return "undefined"; //default
            }
        }

        //create prefab
        public void createPrefab(GameObject obj, string path)
        {
            UnityEngine.Object prefab = PrefabUtility.CreatePrefab(path, obj);
            PrefabUtility.ReplacePrefab(obj, prefab, ReplacePrefabOptions.ConnectToPrefab);
            Debug.Log($"SNUPlugin: Generated prefab in {path}");
        }

        //create directory
        public bool mkdir(string path)
        {
            if (new DirectoryInfo(path).Exists)
            {
                if (!EditorUtility.DisplayDialog("SNUPlugin", Path.GetFileName(path) + "화 프리팹이 이미 존재합니다.\n겹쳐 쓰시겠습니까?", "예", "아니오"))
                {
                    Debug.Log($"SNUPlugin: \"{path}\" already exists. Generation canceled.");
                    return false;
                }
            }
            try
            {
                Directory.CreateDirectory(path);
            } catch
            {
                return false;
            }
            return true;
        }
    }
}
