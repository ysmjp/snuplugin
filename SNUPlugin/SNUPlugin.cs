using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms; //platform-dependant
using System.IO;
using UnityEngine;
using UnityEditor;
using Independentsoft.Office.Spreadsheet;
using System.Text.RegularExpressions;

namespace SNUPlugin
{
    public class SNUPlugin : MonoBehaviour
    {
        private List<Proposal> myProposal = new List<Proposal>();

        void OnStart()
        {
            Debug.Log("SNUPlugin initialized.");
        }

        public bool openDialog()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "기획서 파일을 선택해주세요.";
            ofd.Filter = "스프레드시트 파일(*.xlsx;*.xls)|*.xlsx;*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (!importSheet(ofd.FileName))
                    return false;
                return true;
            }
            return false;
        }

        bool importSheet(string path)
        {
            DobrainGameType dbGameType;
            if (!(new FileInfo(path)).Exists)
                return false;
            //load spreadsheet
            Debug.Log("importSheet: entry");
            Workbook book = new Workbook(path);
            if (book == null)
            {
                EditorUtility.DisplayDialog("SNUPlugin", "엑셀 파일이 아닙니다.", "닫기");
                return false;
            }
            Worksheet sh;
            //create proposal instance
            myProposal.Clear();
            foreach (Sheet sheet in book.Sheets)
            {
                if (sheet is Worksheet)
                    sh = (Worksheet)sheet;
                else
                    continue;
                dbGameType = getGameType(sh);
                if (dbGameType != DobrainGameType.Undefined)
                {
                    myProposal.Add(new Proposal() {
                        filename = path,
                        gametype = dbGameType
                    });
                    Debug.Log(sh.Name + ":" + dbGameType.ToString());
                }
            }
            if (myProposal.Count == 0)
            {
                EditorUtility.DisplayDialog("SNUPlugin", "올바른 기획서 파일이 아닙니다.", "닫기");
                return false;
            }
            return true;
        }

        DobrainGameType getGameType(Worksheet sh)
        {
            bool boolFound = false;
            int targetRow = 0, targetCol = 0;
            string strFormula, strValue;
            DobrainGameType dbGameType;
            //Debug.Log("getGameType: " + sh.Name + ": rows " + sh.Rows.Count + ", cols " + sh.Rows[0].Cells.Count);
            for (int col = 0; col < 3; col++)
            {
                for (int row = 8; row < 13; row++)
                {
                    if (row >= sh.Rows.Count || sh.Rows[row] == null || col >= sh.Rows[row].Cells.Count || sh.Rows[row].Cells[col] == null)
                        continue;
                    strFormula = sh.Rows[row].Cells[col].ToString();
                    strFormula = getTrimmedString(getNormalString(strFormula));
                    //Debug.Log(row + "," + col + ":" + strFormula);
                    if (strFormula == "문제유형")
                    {
                        boolFound = true;
                        targetRow = row;
                        targetCol = col;
                        break;
                    }
                }
                if (boolFound) break;
            }
            if (!boolFound)
                return DobrainGameType.Undefined;
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
                    if (strFormula == "차수")
                        return DobrainGameType.Undefined;
                    if (strFormula == "1" || strFormula == "0")
                        continue;
                    if (row + 1 >= sh.Rows.Count || sh.Rows[row + 1] == null || col >= sh.Rows[row + 1].Cells.Count || sh.Rows[row + 1].Cells[col] == null)
                        continue;
                    strValue = sh.Rows[row + 1].Cells[col].ToString();
                    strValue = getTrimmedString(getNormalString(strValue));
                    //Debug.Log("match: " + (row + 1) + "," + col + ":" + strValue);
                    if (strValue != "1") //true
                        continue;
                    dbGameType = convertToGameType(strFormula);
                    if (dbGameType != DobrainGameType.None)
                        return dbGameType;
                }
            }
            return DobrainGameType.None;
        }

        DobrainGameType convertToGameType(string value)
        {
            switch (value)
            {
                case "":
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

        string getNormalString(string value)
        {
            string strPattern = @"(?=\<)(.*?)(?<=\>)";
            while (Regex.IsMatch(value, strPattern))
                value = Regex.Replace(value, strPattern, "");
            return value;
        }

        string getTrimmedString(string value)
        {
            return value.Replace("\n", "").Replace("\r", "").Replace(" ", "").ToLower();
        }
        
    }

    class Proposal
    {
        public string filename;
        public DobrainGameType gametype;

    }

    //can be generalized by config.json
    enum DobrainGameType
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
}
