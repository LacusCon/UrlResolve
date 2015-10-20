using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Windows;

//EPPlus
using OfficeOpenXml;  //EPPlusクラス
using System.IO;      //FileInfoクラスを使用
using HtmlAgilityPack;
using System.Text.RegularExpressions;

namespace UrlResolve
{
    public class ExcelReslove
    {

        private string excel_inpath = @"Resources\url_set.xlsx";
        //private string template = @"Resources\template.xlsx";

        private List<string> game_url_page = new List<string>();

        private const string _after_name = @"_盗版平台分布";
        private const string _url_name = @"_URL";
        
        private List<url_struct> url_list = new List<url_struct>();
        private List<url_struct> result_list = new List<url_struct>();

        struct url_struct {
            public string sheet_name;
            public string game_name;
            public string platform;
            public string url;
            public string site;
            public string count;
        }

        private Dictionary<string, string> web_key = new Dictionary<string, string>();
        private List<string> platform_list = new List<string>();
        private Dictionary<string, string> regex_key = new Dictionary<string, string>();

        private List<string> errorList = new List<string>();


        public ExcelReslove() {

            platform_list.Add("App Store");//App Store
            platform_list.Add("百度");//百度
            platform_list.Add("腾讯");//腾讯
            platform_list.Add("360");//360
            platform_list.Add("九游");//九游
            platform_list.Add("UC");//UC
            platform_list.Add("PP助手");//PP助手
            platform_list.Add("豌豆荚");//豌豆荚
            platform_list.Add("小米");//小米


            //web_key.Add("App Store", "");//App Store
            //web_key.Add("百度", @"//span[@class='download-num']");//百度
            //web_key.Add("腾讯", @"//div[@class='det-ins-num']");//腾讯
            //web_key.Add("360", @"//span[@class='s-3']");//360[2个]
            //web_key.Add("九游", "");//九游
            //web_key.Add("UC", @"//span[@class='detail-info-down']");//UC  
            //web_key.Add("PP助手", @"//li[@class='borderR']");//PP助手 3
            //web_key.Add("豌豆荚", @"//i[@itemprop='interactionCount']");//豌豆荚 special
            //web_key.Add("小米", @"//div[@class='info-words']");//小米 div info-words(special)
            
            web_key.Add("apple.com", @"");
            web_key.Add("baidu.com", @"//span[@class='download-num']");
            web_key.Add("qq.com", @"//div[@class='det-ins-num']");
            web_key.Add("myapp.com", @"//div[@class='det-ins-num']");
            web_key.Add("u.360.cn", @"//p[@class='g_d_t_nums']");
            web_key.Add("360.cn", @"//span[@class='s-3']");
            web_key.Add("9game.cn", @"");
            web_key.Add("pp.cn", @"//span[@class='detail-info-down']");
            web_key.Add("25pp.com", @"//li[@class='borderR']");
            web_key.Add("wandoujia.com", @"//i[@itemprop='interactionCount']");
            web_key.Add("xiaomi.com", @"//div[@class='info-words']");
            web_key.Add("mi.com", @"");
            
            regex_key.Add("apple.com", @"");
            regex_key.Add("baidu.com", @"^.*: (.*)$");
            regex_key.Add("qq.com", @"^(\d+.*)下载");
            regex_key.Add("myapp.com", @"^(.*)下载");
            regex_key.Add("u.360.cn", @"^.*下载次数：(.*)&nbsp;次");
            regex_key.Add("360.cn", @"^下载：(.*)次");
            regex_key.Add("9game.cn", @"");//no
            regex_key.Add("pp.cn", @"(.*)次下载");
            regex_key.Add("25pp.com", @"^(.*)次下载");
            regex_key.Add("wandoujia.com", @"(.*)");
            regex_key.Add("xiaomi.com", @"^\n.*\n.*\n.*\n.*(\d+.*)次");
            regex_key.Add("mi.com", @"");//no      
        }

        public void ReadExcel()
        {
            LoadExcelData();
        }

        public void LoadExcelData() {

            string sys_path = System.AppDomain.CurrentDomain.BaseDirectory;
            FileInfo fi = new FileInfo(sys_path + excel_inpath);
            using (ExcelPackage excel = new ExcelPackage(fi))
            {
                //処理を記述
                ExcelWorksheets ews = excel.Workbook.Worksheets;
                String temp_name;
                int temp_idx;
                for (int i = 0; i < ews.Count; i++)
                {
                    temp_name = ews[i + 1].Name;
                    temp_idx = temp_name.IndexOf(_url_name);// @"_URL";
                    if (temp_idx > -1)
                    {
                        GetUrl(ews[i + 1], temp_name.Substring(0, temp_idx));
                    }
                }
                //progress_max = url_list.Count;
                MainWindow mw = new MainWindow();

                //process url_list
                for (int i = 0; i < url_list.Count; i++)
                {
                    if (url_list[i].platform.Equals("App Store") 
                        || url_list[i].platform.Equals("九游"))//TODO 忽视apple, 九游
                    {
                        continue;
                    }
                    GetDownCount(url_list[i]);
                    mw.SetProgress(i, url_list.Count);
                }

                //write to excel
                for (int i = 0; i < result_list.Count; i++)
                {
                    
                    SetDownCount(ews, result_list[i]);
                }

                //ews[1].Cells[1, 1].Value = "--------Test---------";
                excel.Save();
                LogOut();
            }
        }

        private void GetUrl(ExcelWorksheet ew, string inSheetName) {

            string _sheetName = inSheetName;
            url_struct _us;

            int rows = ew.Dimension.Rows;// ew.Cells.Rows;
            
            int start_row = 0;
            int start_col = 2;

            start_row = GetDataHead(ew, inSheetName);

            if (start_row == 0)
            {
                return;
            }

            //解析所有url
            for (int i = start_row; i < rows; i++)
            {
                int now_i = i + 1;
                for (int j = 0; j < 9; j++)
                {
                    string url_txt = ew.Cells[now_i, start_col + j + 1].Text;
                    if (String.IsNullOrEmpty(url_txt))
                    {
                        continue;
                    }

                    _us = new url_struct();
                    _us.sheet_name = _sheetName;
                    _us.game_name = ew.Cells[now_i, start_col].Text;
                    _us.platform = platform_list[j];
                    _us.url = ew.Cells[now_i, start_col + j + 1].Text;
                    _us.site = GetSiteName(_us.url);//TODO
                    url_list.Add(_us);
                }
            }

            //Console.WriteLine(ew.Name + "--GetUrl Completed--" + url_list.Count);
        }

        private void GetDownCount(url_struct us)
        {

            bool isExist = false;
            string node_func = "";
            if (web_key.ContainsKey(us.site))
            {
                isExist = true;
                node_func = web_key[us.site];
            }

            if (!isExist || String.IsNullOrEmpty(node_func))
            {
                StringBuilder sb = new StringBuilder("==Error== GetDownCount Sheet::").Append(us.sheet_name).Append("  Plateform::").Append(us.platform);
                errorList.Add(sb.ToString());
                Console.WriteLine(sb.ToString());
                return;
            }
           

            //Console.WriteLine("== GetDownCount ==" + ++test_count);

            HtmlWeb hw = new HtmlWeb();            
            try
            {
                HtmlDocument htmlDoc = hw.Load(us.url);

                HtmlNodeCollection anchors = htmlDoc.DocumentNode.SelectNodes(node_func);//='download-num'
                int aCount = 0;
                string tmp_name = "";

                if (anchors == null || (anchors != null && anchors.Count == 0))
                {
                    StringBuilder sb = new StringBuilder("==Error== GetDownCount  aCount Sheet::").Append(us.sheet_name).Append("  Plateform::").Append(us.platform).Append(" Result:0");
                    errorList.Add(sb.ToString());
                    Console.WriteLine(sb.ToString());
                    return;
                }

                foreach (var item in anchors)
                {
                    aCount++;
                    //tmp_name = item.InnerText;
                    //TODO 加入解析
                    tmp_name = GetResloveText(item.InnerText, GetSiteName(us.url));

                    StringBuilder sb = new StringBuilder("==anchors ==  Sheet::").Append(us.sheet_name).Append(" Game::").Append(us.game_name).Append(" Platform::").Append(us.platform).Append(" URL::").Append(us.url).Append(" Content::").Append(tmp_name);
                    Console.WriteLine(sb.ToString());
                    break;
                }

                if (aCount > 2)//TODO
                {
                    StringBuilder sb = new StringBuilder("==Error== GetDownCount  aCount Sheet::").Append(us.sheet_name).Append("  Plateform::").Append(us.platform).Append(" ##").Append(aCount);
                    errorList.Add(sb.ToString());
                    Console.WriteLine(sb.ToString());
                    return;
                }

                url_struct res_st = new url_struct();
                res_st.sheet_name = us.sheet_name;
                res_st.game_name = us.game_name;
                res_st.platform = us.platform;
                res_st.site = us.site;
                res_st.count = tmp_name;
                result_list.Add(res_st);
            }
            catch (Exception ex)
            {
                StringBuilder sb = new StringBuilder(us.url).Append("== GetDownCount Exception==").Append(ex);
                errorList.Add(sb.ToString());
                Console.WriteLine(sb.ToString());
                throw;
            }

            
        }


        private void SetDownCount(ExcelWorksheets ews, url_struct result_st) {

            //get sheet
            for (int i = 0; i < ews.Count; i++)
            {
                int sheet_idx = i + 1;
                ExcelWorksheet ew = ews[sheet_idx];
                //Console.WriteLine(ews[sheet_idx].Name + "==SetDownCount==" + result_st.sheet_name + _after_name);
                if (ew.Name.Equals(result_st.sheet_name + _after_name))
                {
                    //Console.WriteLine("==SetDownCount Equal== " + ews[sheet_idx].Name);
                    int start_row = 0;

                    //得到起始行
                    start_row = GetDataHead(ew, ew.Name);

                    if (start_row == 0)
                    {
                        return;
                    }

                    //游戏名对照
                    for (int j = start_row; j < ew.Dimension.Rows; j++)
                    {
                        int now_j = j + 1;
                        if (ew.Cells[now_j, 2].Text.Equals(result_st.game_name))
                        {
                            int col_offset = 0;
                            for (int m = 0; m < platform_list.Count; m++)
                            {
                                if (result_st.platform.Equals(platform_list[m]))
                                {
                                    col_offset = m;
                                }
                            }

                            //ew.Cells[now_j, col_offset + 3].Text = result_st.count;//列的偏移需要取
                            ew.Cells[now_j, col_offset + 5].Value = result_st.count;
                            //ews[sheet_idx].Cells[now_j, col_offset + 5].Value = result_st.count;
                            Console.WriteLine("==SetDownCount==" + result_st.game_name + "  Platform::" + result_st.platform + " Content::" + result_st.count);
                            return;
                        }
                    }

                }                
            }
            
        }


        private int GetDataHead(ExcelWorksheet ew, String inSheetName) {

            int rows = ew.Dimension.Rows;// ew.Cells.Rows;
            int start_row = 0;

            //定位开头
            for (int i = 0; i < rows; i++)
            {
                int s_row = i + 1;
                var ecb = ew.Cells[s_row, 1];
                if (ecb.Text.Equals("No."))
                {
                    ecb = ew.Cells[s_row, 2];
                    if (ecb.Text.Equals("游戏"))
                    {
                        start_row = s_row;
                        break;
                    }
                    else
                    {

                        StringBuilder sb = new StringBuilder("==Error== 文档格式不是预定格式 Sheet::").Append(ew.Name).Append("  ==Row(下标起始1)==").Append(i + 1).Append("  --Col--2");
                        errorList.Add(sb.ToString());
                        Console.WriteLine(sb.ToString());
                    }
                }

                if (s_row > 20)//没必要继续
                {
                    StringBuilder sb = new StringBuilder("==Error== 文档格式不是预定格式 Sheet::").Append(ew.Name);
                    errorList.Add(sb.ToString());
                    Console.WriteLine(sb.ToString());
                    return 0;
                }
                //Console.WriteLine(i + "--GetUrl--" + ecb.Text);
            }

            return start_row;
        }


        public string GetResloveText(string inStr, string site) {

            string sRes = inStr;

            try
            {                
                string regex = regex_key[site];

                if (String.IsNullOrEmpty(regex))
                {

                    StringBuilder sb = new StringBuilder("==Error== 不能解析的链接 URL::").Append(site);
                    errorList.Add(sb.ToString());
                    Console.WriteLine(sb.ToString());
                    return sRes;
                }

                if (Regex.IsMatch(inStr, regex))
                {
                    sRes = Regex.Match(inStr, regex).Groups[1].Value;
                }

                sRes = ChangeFormat(sRes);
            }
            catch (Exception ex)
            {
                StringBuilder sb = new StringBuilder("== GetResloveText==").Append(inStr).Append(" ##").Append(ex);
                errorList.Add(sb.ToString());
                Console.WriteLine(sb.ToString());
                throw;
            }


            //有万的去掉，没有的写成0.xxx的形式
            return sRes;
        }

        public string GetPlatformName(int idx) {
            return platform_list[idx];
        }

        public string GetSiteName(string url) {
            string plat_reg = @"^.*?\.(.*?)/";
            string platform = "";
            if (Regex.IsMatch(url, plat_reg))
            {
                platform = Regex.Match(url, plat_reg).Groups[1].Value;
            }

            return platform;
        }

        public string ChangeFormat(string inStr)
        {
            string reg = @"^(.*) *万";
            string resStr = inStr;
            try
            {            
                if (Regex.IsMatch(inStr, reg))
                {
                    resStr = Regex.Match(inStr, reg).Groups[1].Value;
                }
                else {
                    double db = Convert.ToDouble(inStr);
                    resStr = Convert.ToString(db / 10000f);
                }                
            }
            catch (Exception ex)
            {
                StringBuilder sb = new StringBuilder(inStr + "== ChangeFormat ==" + ex);
                errorList.Add(sb.ToString());
                Console.WriteLine(sb.ToString());
                throw;
            }
            return resStr;
        }

        private void LogOut() {

        }


    }
}
