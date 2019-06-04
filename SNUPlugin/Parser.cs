using System;
using System.IO;
using System.Linq;
using UnityEngine;
using UnityEditor;
using Independentsoft.Office.Spreadsheet;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace SNUPlugin
{
    //Excel Parser
    class Parser
    {

        public bool importSheet(List<Proposal> proposalList, string path)
        {
            DobrainGameType dbGameType;
            Proposal prop;
            if (!(new FileInfo(path)).Exists)
                return false;
            //load spreadsheet
            Workbook book = new Workbook(path);
            if (book == null)
            {
                EditorUtility.DisplayDialog("SNUPlugin", "엑셀 파일이 아닙니다.", "닫기");
                return false;
            }
            Worksheet sh;
            //create proposal instance
            proposalList.Clear();
            foreach (Sheet sheet in book.Sheets)
            {
                if (sheet is Worksheet)
                    sh = (Worksheet)sheet;
                else
                    continue;
                dbGameType = getGameType(sh);
                if (dbGameType != DobrainGameType.Undefined)
                {
                    prop = new Proposal()
                    {
                        Filename = path,
                        GameType = dbGameType,
                        DevelopmentType = getDevelopmentType(sh),
                        DevelopmentSubtype = getDevelopmentSubtype(sh),
                        ContentsIndex = getContentsIndex(sh),
                        QuestionIndex = getQuestionIndex(sh)
                    };
                    proposalList.Add(prop);
                    Debug.Log("SNUPlugin: Loaded [" + prop.SimpleFilename + "]" + sh.Name + ":" + prop.GameType.ToString() + "," + prop.DevelopmentType.ToString() + "," + prop.DevelopmentSubtype + "," + prop.ContentsIndex.ToString() + "," + prop.QuestionIndex.ToString());
                }
            }
            if (proposalList.Count == 0)
            {
                EditorUtility.DisplayDialog("SNUPlugin", "올바른 기획서 파일이 아닙니다.", "닫기");
                return false;
            }
            else
            {
                EditorUtility.DisplayDialog("SNUPlugin", proposalList.Count.ToString() + "개의 문제가 로드되었습니다.", "닫기");
            }
            return true;
        }

        public DobrainGameType getGameType(Worksheet sh)
        {
            return findCells(sh, "문제유형", "차수", 8, 0, convertToGameType);
        }

        public DobrainDevelopmentType getDevelopmentType(Worksheet sh)
        {
            return findCells(sh, "항목", "유형", 1, 0, convertToDevelopmentType);
        }

        public string getDevelopmentSubtype(Worksheet sh)
        {
            return findCells(sh, "유형", "공통영역", 4, 0, convertToString);
        }

        public int getContentsIndex(Worksheet sh)
        {
            try
            {
                return (int)(float.Parse(findCells(sh, "콘텐츠번호", "", 0, 0, convertToString, true)));
            }
            catch
            {
                return -1;
            }
        }

        public int getQuestionIndex(Worksheet sh)
        {
            try
            {
                return (int)float.Parse(findCells(sh, "문제번호", "", 0, 3, convertToString, true));
            }
            catch
            {
                return -1;
            }
        }

        public T findCells<T>(Worksheet sh, string matchTitle, string excludedTitle, int startRow, int startCol, Func<string, T> convertFunction)
        {
            return findCells(sh, matchTitle, excludedTitle, startRow, startCol, convertFunction, false);
        }
        public T findCells<T>(Worksheet sh, string matchTitle, string excludedTitle, int startRow, int startCol, Func<string, T> convertFunction, bool justGet)
        {
            bool boolFound = false;
            int targetRow = 0, targetCol = 0;
            string strFormula, strValue;
            T res;
            //Debug.Log("getGameType: " + sh.Name + ": rows " + sh.Rows.Count + ", cols " + sh.Rows[0].Cells.Count);
            for (int col = startCol; col < startCol + 3; col++)
            {
                for (int row = startRow; row < startRow + 5; row++)
                {
                    if (row >= sh.Rows.Count || sh.Rows[row] == null || col >= sh.Rows[row].Cells.Count || sh.Rows[row].Cells[col] == null)
                        continue;
                    strFormula = sh.Rows[row].Cells[col].ToString();
                    strFormula = getTrimmedString(getNormalString(strFormula));
                    //Debug.Log(row + "," + col + ":" + strFormula);
                    if (strFormula == matchTitle)
                    {
                        if (justGet)
                        {
                            if (row + 1 >= sh.Rows.Count || sh.Rows[row + 1] == null || col >= sh.Rows[row + 1].Cells.Count || sh.Rows[row + 1].Cells[col] == null)
                                return convertFunction("undefined");
                            strFormula = sh.Rows[row + 1].Cells[col].ToString();
                            strFormula = getTrimmedString(getNormalString(strFormula));
                            //Debug.Log(strFormula);
                            return convertFunction(strFormula);
                        }
                        else
                        {
                            boolFound = true;
                            targetRow = row;
                            targetCol = col;
                            break;
                        }
                    }
                }
                if (boolFound) break;
            }
            if (!boolFound)
                return convertFunction("undefined");
            //Debug.Log("getGameType: found offset " + targetRow + "," + targetCol);
            for (int row = targetRow + 1; row < targetRow + 6; row++)
            {
                for (int col = targetCol; col < targetCol + 20; col++)
                {
                    if (row >= sh.Rows.Count || sh.Rows[row] == null || col >= sh.Rows[row].Cells.Count || sh.Rows[row].Cells[col] == null)
                        continue;
                    strFormula = sh.Rows[row].Cells[col].ToString();
                    strFormula = getTrimmedString(getNormalString(strFormula));
                    //Debug.Log(row + "," + col + ":" + strFormula);
                    if (strFormula == excludedTitle)
                        return convertFunction("none");
                    if (strFormula == "1" || strFormula == "0")
                        continue;
                    if (row + 1 >= sh.Rows.Count || sh.Rows[row + 1] == null || col >= sh.Rows[row + 1].Cells.Count || sh.Rows[row + 1].Cells[col] == null)
                        continue;
                    strValue = sh.Rows[row + 1].Cells[col].ToString();
                    strValue = getTrimmedString(getNormalString(strValue));
                    //Debug.Log("match: " + (row + 1) + "," + col + ":" + strValue);
                    if (strValue != "1") //true
                        continue;
                    res = convertFunction(strFormula);
                    if (!res.Equals(convertFunction("none")))
                        return res;
                }
            }
            return convertFunction("none");
        }
        public DobrainGameType convertToGameType(string value)
        {
            switch (value)
            {
                case "undefined":
                    return DobrainGameType.Undefined;
                case "none":
                    return DobrainGameType.None;
                case "다지선일":
                    return DobrainGameType.ChoiceOne;
                case "움직이는두더지게임":
                    return DobrainGameType.MoleMoving;
                case "움직이지않는두더지게임":
                    return DobrainGameType.MoleStatic;
                case "애니메이팅두더지게임":
                    return DobrainGameType.MoleAnimating;
                case "그림그리기":
                    return DobrainGameType.DrawingImage;
                case "패턴그리기":
                    return DobrainGameType.DrawingPattern;
                case "드래그앤드랍":
                    return DobrainGameType.DragAndDrop;
                case "다지선다":
                    return DobrainGameType.ChoiceMulti;
                case "지우기게임":
                    return DobrainGameType.Erase;
                case "발견하기문제(틀린그림,발자국)":
                    return DobrainGameType.Catch;
                case "선긋기":
                    return DobrainGameType.DrawingLine;
                case "순서대로누르기":
                    return DobrainGameType.Click;
                case "블록쌓기":
                    return DobrainGameType.PilingBlocks;
                case "카드뒤집기":
                    return DobrainGameType.FlipingCards;
                case "기타":
                    return DobrainGameType.Other;
                default:
                    return DobrainGameType.None;
            }
        }
        public DobrainDevelopmentType convertToDevelopmentType(string value)
        {
            switch (value)
            {
                case "undefined":
                    return DobrainDevelopmentType.Undefined;
                case "none":
                    return DobrainDevelopmentType.None;
                case "지각속도력":
                    return DobrainDevelopmentType.PerceptionSpeed;
                case "공간지각력":
                    return DobrainDevelopmentType.SpatialPerception;
                case "수리력":
                    return DobrainDevelopmentType.Mathematical;
                case "창의력":
                    return DobrainDevelopmentType.Creative;
                case "시각적변별력":
                    return DobrainDevelopmentType.VisualDiscrimination;
                case "기억력":
                    return DobrainDevelopmentType.Memory;
                case "구성력":
                    return DobrainDevelopmentType.Compositive;
                case "추론력":
                    return DobrainDevelopmentType.Inferential;
                default:
                    return DobrainDevelopmentType.None;
            }

        }
        public string convertToString(string value)
        {
            return value;
        }
        public string getNormalString(string value)
        {
            string strPattern = @"(?=\<)(.*?)(?<=\>)";
            while (Regex.IsMatch(value, strPattern))
                value = Regex.Replace(value, strPattern, "");
            return value;
        }
        public string getTrimmedString(string value)
        {
            return value.Replace("\n", "").Replace("\r", "").Replace(" ", "").ToLower();
        }
    }

    //Proposal Structure
    public class Proposal
    {
        public string Filename
        {
            get { return _Filename; }
            set
            {
                _Filename = value;
                SimpleFilename = Path.GetFileName(value);
            }
        } //파일 주소
        private string _Filename; //inner memeber
        public string SimpleFilename; //파일 이름
        public int ContentsIndex = 0; //컨텐츠 번호
        public int QuestionIndex = 0; //문제 번호
        public DobrainDevelopmentType DevelopmentType; //항목
        public string DevelopmentSubtype = ""; //유형
        public DobrainGameType GameType; //문제 유형
    }

    //can be generalized by config.json
    public enum DobrainGameType
    {
        Undefined, //have no gametype cell
        None, //none of gametypes selected
        ChoiceOne, //다지선일,
        MoleMoving, //움직이는 두더지게임,
        MoleStatic, //움직이지 않는 두더지게임,
        MoleAnimating, //애니메이팅 두더지 게임,
        DrawingImage, //그림그리기,
        DrawingPattern, //패턴 그리기,
        DragAndDrop, //드래그앤드랍,
        ChoiceMulti, //다지선다,
        Erase, //지우기게임,
        Catch, //발견하기문제(틀린그림, 발자국),
        DrawingLine, //선긋기,
        Click, //순서대로누르기,
        PilingBlocks, //블록쌓기,
        FlipingCards, //카드뒤집기,
        Other, //기타
    }

    public enum DobrainDevelopmentType
    {
        Undefined,
        None,
        PerceptionSpeed, //지각속도력
        SpatialPerception, //공간지각력
        Mathematical, //수리력
        Creative, //창의력
        VisualDiscrimination,//시각적변별력
        Memory, //기억력
        Compositive, //구성력
        Inferential //추론력
    }
}
