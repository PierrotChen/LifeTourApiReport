using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Configuration;
using System.IO;
using System.Drawing;



namespace AGT_EmpGrupReport
{

    internal class AgtEmpReport
    {
        private ExcelPackage ep;
        private ExcelWorksheet sheet;
        private List<string> ProdCategory;
        private List<YearMonth> YearMonth;
        private List<EmpList> GroupType, EmpList;
        private int rows, cells, Qt;
        private int Gsort;
        private ExcelRange range;

        internal string Go(List<TABLE> agtData)
        {
            string _filePath = CreateFile();
            GetGroupType(agtData);
            foreach (var _yearMonth in YearMonth)
            {
                SetFormat(_yearMonth);
                SetLogic(_yearMonth, agtData);
                SetBorder();
            }
            CreateRuleSheet();
            ep.Save();
            return _filePath;
        }

        private string CreateFile()
        {
            string Drive;
            Drive = (ConfigurationManager.AppSettings["Drive"]);
            string dt = DateTime.Now.ToString("yyyyMMdd");
            FileInfo filePath = new FileInfo(Drive + dt + "全省同業部團體業績管理報表.xlsx");
            if (File.Exists(filePath.ToString()))
            {
                File.Delete(filePath.ToString());
            }
            ep = new ExcelPackage(filePath);
            return filePath.ToString();
        } //建立檔案
        private void GetGroupType(List<TABLE> agtData)
        {
            //↓↓各組別指標 Group-Type//
            GroupType = new List<AGT_EmpGrupReport.EmpList>();

            var temp = agtData.Where(x => x.Region != "").Select(x => new { _groupType = x.GroupType, _region = x.Region }).Distinct().OrderBy(x => x._groupType).ToList();
            //GroupType = agtData.OrderBy(x => x.GroupType).Select(x => x.GroupType).Distinct().ToList();

            int keepGroupType = 0;
            foreach (var _temp in temp)
            {
                if (_temp._groupType != keepGroupType)
                {
                    GroupType.Add(new AGT_EmpGrupReport.EmpList { Region = _temp._region, GroupType = _temp._groupType });
                    keepGroupType = _temp._groupType;
                }
                
            }
            //↓↓線別 & 排序//
            ProdCategory = new List<string>();
            ProdCategory = agtData.OrderBy(x => x.PSort).Where(x => x.Prod_Category != "").Select(x => x.Prod_Category).Distinct().ToList();
            //↓↓各月份 - 對應時間頁籤 //
            var ymResult = agtData.Where(x => x.LeavM != 0).Select(x => new { LeavM = x.LeavM, LeavY = x.LeavY }).Distinct().OrderBy(x => x.LeavM).OrderBy(x => x.LeavY).ToList();
            YearMonth = new List<AGT_EmpGrupReport.YearMonth>();
            YearMonth = ymResult.Select(x => new YearMonth { LeavM = x.LeavM, LeavY = x.LeavY }).ToList();

            //↓↓業務分組
            var EmpWithGroup = agtData.Select(x => new
            { x.EMP_CNM, x.GroupSort, x.GroupType, x.Region }

            ).Distinct().OrderBy(x => x.GroupSort).OrderBy(x => x.GroupType).ToList();

            EmpList = new List<AGT_EmpGrupReport.EmpList>();
            foreach (var _groupType in GroupType)
            {
                foreach (var _empWithGroup in EmpWithGroup)
                {
                    if (_groupType.GroupType == _empWithGroup.GroupType)
                    {
                        EmpList.Add(new EmpList
                        {
                            Emp_cnm = _empWithGroup.EMP_CNM,
                            GroupSort = _empWithGroup.GroupSort,
                            GroupType = _empWithGroup.GroupType,
                            Region = _empWithGroup.Region,
                        });
                        Gsort = _empWithGroup.GroupSort + 1;
                    }
                }
                string[] leftTitle = new string[] { "團小計", "票小計" };

                foreach (var _leftTitle in leftTitle)
                {
                    EmpList.Add(new EmpList
                    {
                        Emp_cnm = _leftTitle,
                        GroupType = _groupType.GroupType,
                        GroupSort = Gsort,
                    }
                    );
                    Gsort++;
                }
            }
            Console.WriteLine("");
        } //分類 & 取得所需部件
        private void SetFormat(YearMonth _yearMonth)
        {

            rows = 3; cells = 2;
            int TopCells = 3, TopRows = 2;
            string sheetName = _yearMonth.LeavY + "年" + _yearMonth.LeavM + "月";
            sheet = ep.Workbook.Worksheets.Add(sheetName);
            sheet.Cells[1, 1].Value = "全省 北中南區" + sheetName;
            foreach (var _prodCategory in ProdCategory)
            {
                sheet.Cells[TopRows, TopCells].Value = _prodCategory;
                TopCells++;
            }
            sheet.Cells[1, 1, 1, TopCells - 1].Merge = true;

            foreach (var _type in GroupType)
            {
                int _tempRows = rows;
                //sheet.Cells[rows, 1].Value = "北" + _type + "課";
                sheet.Cells[rows, 1].Value = _type.Region;
                foreach (var _empList in EmpList)
                {
                    if (_type.GroupType == _empList.GroupType)
                    {
                        //if (_empList.GroupSort == 1)
                        //{
                        //    //sheet.Cells[rows, cells].Value = "北" + _empList.GroupType + "課目標";
                        //    //rows++;
                        //    sheet.Cells[rows, cells].Value = _empList.Emp_cnm;
                        //}
                        //else if (_empList.Emp_cnm == "")
                        if (_empList.Emp_cnm == "")
                        {
                            sheet.Cells[rows, cells].Value = "待補";
                        }
                        else
                        {
                            sheet.Cells[rows, cells].Value = _empList.Emp_cnm;
                        }
                        rows++;
                    }
                }

                sheet.Cells[_tempRows, 1, rows - 1, 1].Merge = true;



            }


        } //設定Excel表格格式
        private void SetLogic(YearMonth _yearMonth, List<TABLE> agtData)
        {
            cells = 3;
            foreach (var _prod in ProdCategory)
            {
                rows = 3;
                foreach (var _empList in EmpList)
                {
                    Qt = 0;
                    if (_empList.Emp_cnm == "團小計")
                    {
                        foreach (var _agtData in agtData)
                        {
                            if (_agtData.GroupType == _empList.GroupType && _prod == _agtData.Prod_Category && _yearMonth.LeavY == _agtData.LeavY && _yearMonth.LeavM == _agtData.LeavM)
                            {
                                Qt += _agtData.Pax_Qt;
                                //Pax_qt += _agtData.Pax_Qt;
                            }
                        }
                    }
                    else if (_empList.Emp_cnm == "票小計")
                    {
                        foreach (var _agtData in agtData)
                        {
                            if (_agtData.GroupType == _empList.GroupType && _prod == _agtData.Prod_Category && _yearMonth.LeavY == _agtData.LeavY && _yearMonth.LeavM == _agtData.LeavM)
                            {
                                Qt += _agtData.Tkt_Qt;
                            }
                        }
                    }
                    else
                    {
                        foreach (var _agtData in agtData)
                        {
                            if (_agtData.EMP_CNM == _empList.Emp_cnm && _prod == _agtData.Prod_Category && _yearMonth.LeavY == _agtData.LeavY && _yearMonth.LeavM == _agtData.LeavM)
                            {
                                Qt += _agtData.Pax_Qt;
                            }
                        }
                    }
                    sheet.Cells[rows, cells].Value = Qt;
                    rows++;
                }
                cells++;
            }


        }  //設定邏輯 & 寫值
        private void SetBorder()
        {

            int startRowNumber = sheet.Dimension.Start.Row;//起始列編號，從1算起
            int endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
            int startColumn = sheet.Dimension.Start.Column;//開始欄編號，從1算起
            int endColumn = sheet.Dimension.End.Column;//結束欄編號，從1算起

            for (int i = 1; i < endRowNumber + 1; i++)
            {
                for (int j = 1; j < endColumn + 1; j++)
                {
                    sheet.Cells[i, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); //儲存格框線
                    sheet.Column(j).AutoFit();
                }
            }

            range = sheet.Cells[startRowNumber, startColumn, endRowNumber, endColumn];
            range.Style.Numberformat.Format = "##,##0";
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.White);
            range.Style.Font.Name = "微軟正黑體";
            range.Style.Font.Color.SetColor(Color.Black);
            range.Style.Font.Size = 12;
            sheet.Row(1).Style.Font.Size = 18;
            sheet.Column(1).Style.Font.Size = 12;
            sheet.Column(1).Width = 10;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        } //給予Excel格線
        private void CreateRuleSheet()
        {
            sheet = ep.Workbook.Worksheets.Add("規則 & 定義");
            string[] RuleText = new string[]{
        "",
        "全省同業部 北中南區 團體業績管理報表 規則定義如下：",
        "",
        "",
        "一、時間區間：",
        "",
        "       依團體出發日，取資料產出日之前一月、後三月資料",
        "",
        "二、人數：",
        "",
        "       訂單狀態為「非取消」且",
        "       參團狀態為「確認」且",
        "       參團類別為「全程」以及「TKTonly」(不含JoinTour)之人頭數",
        "",
        "三、部門：",
        "",
        "北區同業事業群-北區 1-4 課(台北)、5課(中壢)、6課(新竹)",
        "中區同業事業群-中區 1、2課(台中)",
        "南區同業事業群-南區 1課(高雄)、南區2課(台南、嘉義)",
        "",
        "",
            };

            for (int i = 1; i < RuleText.Count(); i++)
            {
                sheet.Cells[i, 1].Value = RuleText[i];
                sheet.Cells[i, 1].Style.Font.Name = "微軟正黑體";
                sheet.Cells[i, 1].Style.Font.Color.SetColor(Color.Black);
                sheet.Cells[i, 1].Style.Font.Size = 12;
                sheet.Row(1).Style.Font.Size = 18;
            }
            sheet.Column(1).AutoFit();
        } //規則頁籤
    }
}