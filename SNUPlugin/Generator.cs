using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEngine;
using UnityEngine.Events;
using UnityEngine.UI;
using UnityEditor;
using UnityEditor.Events;

namespace SNUPlugin
{
    //prefab generator
    class Generator
    {
        //generate prefab
        public bool generate(SNUPlugin snuplugin, List<Proposal> proposalList)
        {
            string strDir, strPath, strObjectName, strQManagerName;
            int intSuccess = 0;
            Proposal prop;
            GameObject objQuestion = null, objQManager = null, objExplanation = null, objExample = null, objElement = null;
            List<GameObject> lstObjExample = new List<GameObject>(), lstObjElement = new List<GameObject>();
            GameObject[] objStep = new GameObject[6];

            // Design Constants
            int BoardWidth = 1030, BoardHeight = 740;

            if (proposalList.Count == 0) return false;
            strDir = "Assets/Contents/Resources/Weekday/" + proposalList[0].ContentsIndex.ToString();
            if (!mkdir(strDir)) return false; //fail to make direcotry
            for (int i = 0; i < proposalList.Count; i++) //loop thru proposal list
            {
                prop = proposalList[i];
                strObjectName = "ch" + prop.ContentsIndex.ToString() + "_" + prop.QuestionIndex.ToString();
                strPath = strDir + "/" + strObjectName + ".prefab";
                strQManagerName = getQManagerName(prop);
                if (strQManagerName == "undefined")
                {
                    objQuestion = new GameObject(strObjectName);
                }
                else
                {
                    objQuestion = new GameObject(strObjectName); //ch#_#
                    objQuestion.AddComponent<RectTransform>();

                    objQManager = new GameObject(strQManagerName); //*QManager
                    objQManager.transform.SetParent(objQuestion.transform);
                    objQManager.AddComponent<RectTransform>();
                    AddComponentExt(objQManager, strQManagerName);

                    objExplanation = new GameObject("Explanation Board");
                    objExplanation.transform.SetParent(objQManager.transform);
                    objExplanation.AddComponent<RectTransform>();
                    objExplanation.GetComponent<RectTransform>().sizeDelta = new Vector2(BoardWidth, BoardHeight);
                    objExplanation.AddComponent<CanvasRenderer>();
                    AddComponentExt(objExplanation, "Image");
                    int nStep = 6;
                    int nChoice = 5; // We would get this from proposal sheet
                    // Create Steps
                    for (int j = 0; j < nStep; j++)
                    {
                        objStep[j] = new GameObject("Step" + (j > 0 ? $" ({j})" : ""));
                        objStep[j].transform.SetParent(objQManager.transform);
                        objStep[j].AddComponent<RectTransform>();
                        AddComponentExt(objStep[j], "DerivedQuestion");
                        objStep[j].AddComponent<Animator>();
                        objStep[j].SetActive(true);

                        objExample = new GameObject("Example");
                        objExample.transform.SetParent(objStep[j].transform);
                        objExample.AddComponent<RectTransform>();
                        lstObjExample.Add(objExample);

                        Debug.Log("Snuplugin works well " + objQManager.name);
                        switch (objQManager.name) {
                            case "ClickOneQManager":
                                for (int k = 0; k < nChoice; k++) {
                                    string strName = "Element" + k.ToString();
                                    GameObject objChoice = new GameObject(strName);
                                    objChoice.transform.SetParent(objStep[j].transform);
                                    objChoice.AddComponent<RectTransform>();
                                    objChoice.GetComponent<RectTransform>().sizeDelta = new Vector2(50, 50);
                                    objChoice.transform.position = new Vector3((k-nChoice/2)*(BoardWidth/nChoice - 30), 0, 0);
                                    Button b = objChoice.AddComponent<Button>();
                                    
                                    Debug.Log(b.isActiveAndEnabled.ToString());
                                    var abcd = new UnityEngine.Events.UnityAction(() => Debug.Log("FUCK"));
                                    var targetinfo = UnityEvent.GetValidMethodInfo(this, "OnButtonClick", new Type[] { typeof(GameObject)});
                                    UnityAction<GameObject> action = Delegate.CreateDelegate(typeof(UnityAction<GameObject>), this, targetinfo, false) as UnityAction<GameObject>;

                                    var d = System.Delegate.CreateDelegate(typeof(UnityAction), this, "OnButtonClick") as UnityAction;

                                    UnityEventTools.AddObjectPersistentListener(b.onClick, action, objChoice);

                                    b.onClick.Invoke();
                                }
                                break;
                            case "DragAndDropQManager":
                                break;
                        }


                        // General Objects
                        /*
                        objElement = new GameObject("Element");
                        objElement.transform.SetParent(objStep[j].transform);
                        objElement.AddComponent<RectTransform>();
                        lstObjElement.Add(objElement);

                        objElement = new GameObject("Element (1)");
                        objElement.transform.SetParent(objStep[j].transform);
                        objElement.AddComponent<RectTransform>();
                        lstObjElement.Add(objElement);
                        */
                    }

                }
                // enable first step only
                //objStep[0].SetActive(true);

                createPrefab(objQuestion, strPath); //create and save prefab

                for (int j = 0; j < 6; j++)
                        snuplugin.destroyObject(objStep[j]);
                foreach (GameObject obj in lstObjExample)
                    snuplugin.destroyObject(obj);
                foreach (GameObject obj in lstObjElement)
                    snuplugin.destroyObject(obj);
                snuplugin.destroyObject(objExplanation);
                snuplugin.destroyObject(objQManager);
                snuplugin.destroyObject(objQuestion);
                               
                intSuccess++;
            }
            EditorUtility.DisplayDialog("SNUPlugin", $"{Path.GetFileName(proposalList[0].Filename)}:\n{strDir}에\n{proposalList[0].ContentsIndex.ToString()}번 컨텐츠의 Prefab {intSuccess}개를 생성하였습니다.", "닫기");
            return true;
        }

        public void OnButtonClick() {
            Debug.Log("Button Clicked");
            return;
        }

        public Component AddComponentExt(GameObject obj, string scriptName)
        {
            Component cmpnt = null;


            for (int i = 0; i < 10; i++)
            {
                //If call is null, make another call
                cmpnt = _AddComponentExt(obj, scriptName, i);

                //Exit if we are successful
                if (cmpnt != null)
                {
                    break;
                }
            }


            //If still null then let user know an exception
            if (cmpnt == null)
            {
                Debug.LogError("Failed to Add Component");
                return null;
            }
            return cmpnt;
        }

        private Component _AddComponentExt(GameObject obj, string className, int trials)
        {
            //Any script created by user(you)
            const string userMadeScript = "Assembly-CSharp, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null";
            //Any script/component that comes with Unity such as "Rigidbody"
            const string builtInScript = "UnityEngine, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null";

            //Any script/component that comes with Unity such as "Image"
            const string builtInScriptUI = "UnityEngine.UI, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null";

            //Any script/component that comes with Unity such as "Networking"
            const string builtInScriptNetwork = "UnityEngine.Networking, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null";

            //Any script/component that comes with Unity such as "AnalyticsTracker"
            const string builtInScriptAnalytics = "UnityEngine.Analytics, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null";

            //Any script/component that comes with Unity such as "AnalyticsTracker"
            const string builtInScriptHoloLens = "UnityEngine.HoloLens, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null";

            Assembly asm = null;

            try
            {
                //Decide if to get user script or built-in component
                switch (trials)
                {
                    case 0:

                        asm = Assembly.Load(userMadeScript);
                        break;

                    case 1:
                        //Get UnityEngine.Component Typical component format
                        className = "UnityEngine." + className;
                        asm = Assembly.Load(builtInScript);
                        break;
                    case 2:
                        //Get UnityEngine.Component UI format
                        className = "UnityEngine.UI." + className;
                        asm = Assembly.Load(builtInScriptUI);
                        break;

                    case 3:
                        //Get UnityEngine.Component Video format
                        className = "UnityEngine.Video." + className;
                        asm = Assembly.Load(builtInScript);
                        break;

                    case 4:
                        //Get UnityEngine.Component Networking format
                        className = "UnityEngine.Networking." + className;
                        asm = Assembly.Load(builtInScriptNetwork);
                        break;
                    case 5:
                        //Get UnityEngine.Component Analytics format
                        className = "UnityEngine.Analytics." + className;
                        asm = Assembly.Load(builtInScriptAnalytics);
                        break;

                    case 6:
                        //Get UnityEngine.Component EventSystems format
                        className = "UnityEngine.EventSystems." + className;
                        asm = Assembly.Load(builtInScriptUI);
                        break;

                    case 7:
                        //Get UnityEngine.Component Audio format
                        className = "UnityEngine.Audio." + className;
                        asm = Assembly.Load(builtInScriptHoloLens);
                        break;

                    case 8:
                        //Get UnityEngine.Component SpatialMapping format
                        className = "UnityEngine.VR.WSA." + className;
                        asm = Assembly.Load(builtInScriptHoloLens);
                        break;

                    case 9:
                        //Get UnityEngine.Component AI format
                        className = "UnityEngine.AI." + className;
                        asm = Assembly.Load(builtInScript);
                        break;
                }
            }
            catch
            {
                //Debug.Log("Failed to Load Assembly" + e.Message);
            }

            //Return if Assembly is null
            if (asm == null)
            {
                return null;
            }

            //Get type then return if it is null
            Type type = asm.GetType(className);
            if (type == null)
                return null;

            //Finally Add component since nothing is null
            Component cmpnt = obj.AddComponent(type);
            return cmpnt;
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
                    return "QuestionManager";//default
                case DobrainGameType.None:
                    return "QuestionManager";//default
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
                    return "QuestionManager";//default
                default:
                    return "QuestionManager"; //default
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
