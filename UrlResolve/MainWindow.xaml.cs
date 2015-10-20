using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using HtmlAgilityPack;

namespace UrlResolve
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }



        private void UrlResolve_Click(object sender, RoutedEventArgs e)
        {
            ExcelReslove er = new ExcelReslove();
            er.ReadExcel();

            //CheckTest();

            //TestChg();

            MessageBox.Show("== Mission Completed ==");
        }

        private void TestChg() {
            ExcelReslove er = new ExcelReslove();
            List<string> lst = new List<string>();
            lst.Add("4789");
            lst.Add("876");
            lst.Add("87.12万");
            lst.Add("89.22 万");

            for (int i = 0; i < lst.Count; i++)
            {
                string cf = er.ChangeFormat(lst[i]);
                Console.WriteLine("==TestChg==" + cf);
            }
        }


        private void CheckTest() {

            HtmlWeb hw = new HtmlWeb();
        
            List<string> url_list = new List<string>();
            url_list.Add("");//Apple
            url_list.Add(@"http://shouji.baidu.com/game/item?docid=6782660&from=landing&f=search_app_%E9%93%B6%E9%AD%82%40list_4_title%404%40");//百度
            url_list.Add(@"http://android.myapp.com/myapp/detail.htm?apkName=net.cyyun.jumpball_gintama");//腾讯
            url_list.Add(@"http://zhushou.360.cn/detail/index/soft_id/95808?recrefer=SE_D_%E9%93%B6%E9%AD%82%E6%B6%88%E6%B6%88%E7%9C%8B");//360[2个]
            url_list.Add("");//九游
            url_list.Add(@"http://m.pp.cn/detail.html?query=NBA%E5%85%A8%E6%98%8E%E6%98%9F%E6%8C%91%E6%88%98%E8%B5%9B&ch=uc&ch_src=sm&appid=228685");//UC  
            url_list.Add(@"http://android.25pp.com/detail_199009.html");//PP助手 3
            url_list.Add(@"http://www.wandoujia.com/apps/com.hotheadgames.google.free.bigwinbasketball");//豌豆荚 special
            url_list.Add(@"http://game.xiaomi.com/app-appdetail--app_id__5327.html");//小米 div info-words(special)
            url_list.Add(@"http://ku.u.360.cn/detail.php?s=web&sid=70916");//360-2
            url_list.Add(@"http://sj.qq.com/myapp/detail.htm?apkName=com.freeverse.nbas");//qq-2            


            List<string> key_lst = new List<string>();
            key_lst.Add("");//Apple
            key_lst.Add(@"//span[@class='download-num']");//百度
            key_lst.Add(@"//div[@class='det-ins-num']");//腾讯
            key_lst.Add(@"//span[@class='s-3']");//360[2个]
            key_lst.Add("");//九游
            key_lst.Add(@"//span[@class='detail-info-down']");//UC  TODO
            key_lst.Add(@"//li[@class='borderR']");//PP助手 3
            key_lst.Add(@"//i[@itemprop='interactionCount']");//豌豆荚 special
            key_lst.Add(@"//div[@class='info-words']");//小米 div info-words(special)
            key_lst.Add(@"//p[@class='g_d_t_nums']");//360-2
            key_lst.Add(@"//div[@class='det-ins-num']");//qq-2

            int idx = 2;
            try
            {
                HtmlDocument htmlDoc = hw.Load(url_list[idx]);
                HtmlNodeCollection anchors = htmlDoc.DocumentNode.SelectNodes(key_lst[idx]);//='download-num'
                if (anchors == null)
                {
                    return;
                }
                foreach (var item in anchors)
                {
                    ExcelReslove er = new ExcelReslove();
                    string res = er.GetResloveText(item.InnerText, er.GetSiteName(url_list[idx]));
                    Console.WriteLine(item.InnerText + "==anchors==" + res);
                    //break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("--Test--" + ex);
                throw;
            }
        }

        public void SetProgress(int progress, int max)
        {
            UrlResolveBtn.Content = new StringBuilder("当前处理进度：").Append(progress).Append(" / ").Append(max);
        }

    }
}
