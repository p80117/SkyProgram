using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;    //add
using Excel = Microsoft.Office.Interop.Excel;   // 切記要使用Excel Library
using System.Threading;

namespace Netlist_Compare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConstString_Path = "";

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ConstString_Path = dialog.FileName;
                textBox_File_Name1.Text = ConstString_Path;
                //MessageBox.Show(ConstString_Path);

            }


            string ConstString_Path_Old = textBox_File_Name1.Text;
            if (ConstString_Path_Old != "")
            {
                string[] lines = File.ReadAllLines(ConstString_Path_Old);
                string[] Item_Value1;
                string[] Item_Value2;
                int package_net_bit = 0;
                string temp_value;
                for (int i = 0; i < lines.Length; i++)
                {

                    temp_value = lines[i].ToString();

                    switch (temp_value)
                    {
                        case "$NETS":
                        case "$END":
                            package_net_bit++;
                            break;

                        default:

                            if (package_net_bit == 1)
                            {
                                Item_Value1 = lines[i].ToString().Split(';');

                                string[] row = new string[Item_Value1.Length];

                                //row[0] = Item_Value1[0].Trim();

                                for (int j = 0; j < Item_Value1.Length; j++)
                                {
                                    row[j] = Item_Value1[j].Trim();
                                }

                                if (Item_Value1.Length == 1)
                                {
                                    Item_Value2 = Item_Value1[0].ToString().Split(',');   //check next value is different net
                                    Item_Value2 = Item_Value2[0].ToString().Split(' ');   //check next value is different net

                                    int aaa = table_net.Rows.Count;
                                    string ccc;
                                    for (int j = 0; j < Item_Value2.Length; j++)
                                    {
                                        if (Item_Value2[j].Trim() != "")
                                        {
                                            ccc = table_net.Rows[aaa - 1][1].ToString() + " " + Item_Value2[j].Trim();
                                            table_net.Rows[aaa - 1][1] = ccc;
                                        }
                                    }
                                }
                                else
                                {
                                    Item_Value2 = row[1].ToString().Split(',');   //check next value is different net

                                    if (Item_Value2.Length == 2)
                                    {
                                        int aaa = table_net.Rows.Count;
                                        //int bbb = table_net.Columns.Count;

                                        //string ccc = table_net.Rows[aaa][1].ToString();

                                        row[1] = Item_Value2[0].Trim();

                                    }
                                    table_net.Rows.Add(row);
                                }
                            }

                            if ((temp_value != "$PACKAGES") && package_net_bit == 0)
                            {
                                Item_Value1 = lines[i].ToString().Split('!');
                                Item_Value2 = Item_Value1[1].ToString().Split(';');
                                string[] row = new string[Item_Value1.Length + Item_Value2.Length - 1];


                                row[0] = Item_Value1[0].Trim();

                                for (int j = 0; j < Item_Value2.Length; j++)
                                {
                                    row[j + 1] = Item_Value2[j].Trim();
                                }
                                table.Rows.Add(row);
                            }
                            break;
                    }
                }
                //string[] lines = File.ReadAllLines(@"file path");
            }
            else
                MessageBox.Show("InputFileError");

        }

        DataTable table = new DataTable();
        DataTable table_net = new DataTable();
        DataTable table_new = new DataTable();
        DataTable table_net_new = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            //--Set_Title_Name--------------------------//
            table.Columns.Add("Pacakge", typeof(string));
            table.Columns.Add("Part Value", typeof(string));
            table.Columns.Add("Location", typeof(string));
            //table.Columns.Add("Age", typeof(string));
            //------------------------------------------//
            dataGridView1.DataSource = table;

            //--Set_Title_Name--------------------------//
            table_net.Columns.Add("Net Name", typeof(string));
            table_net.Columns.Add("Netlist", typeof(string));
            //table.Columns.Add("Age", typeof(string));
            //------------------------------------------//
            dataGridView2.DataSource = table_net;

            //--Set_Title_Name--------------------------//
            table_new.Columns.Add("Pacakge", typeof(string));
            table_new.Columns.Add("Part Value", typeof(string));
            table_new.Columns.Add("Location", typeof(string));
            //table.Columns.Add("Age", typeof(string));
            //------------------------------------------//
            dataGridView3.DataSource = table_new;

            //--Set_Title_Name--------------------------//
            table_net_new.Columns.Add("Net Name", typeof(string));
            table_net_new.Columns.Add("Netlist", typeof(string));
            //table.Columns.Add("Age", typeof(string));
            //------------------------------------------//
            dataGridView4.DataSource = table_net_new;


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string ConstString_Path = "";

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ConstString_Path = dialog.FileName;
                textBox_File_Name2.Text = ConstString_Path;
                //MessageBox.Show(ConstString_Path);

            }
            string ConstString_Path_New = textBox_File_Name2.Text;
            if (ConstString_Path_New != "")
            {
                string[] lines = File.ReadAllLines(ConstString_Path_New);
                string[] Item_Value1;
                string[] Item_Value2;
                int package_net_bit = 0;
                string temp_value;
                for (int i = 0; i < lines.Length; i++)
                {

                    temp_value = lines[i].ToString();

                    switch (temp_value)
                    {
                        case "$NETS":
                        case "$END":
                            package_net_bit++;
                            break;

                        default:

                            if (package_net_bit == 1)
                            {
                                Item_Value1 = lines[i].ToString().Split(';');

                                string[] row = new string[Item_Value1.Length];

                                for (int j = 0; j < Item_Value1.Length; j++)
                                {
                                    row[j] = Item_Value1[j].Trim();
                                }

                                if (Item_Value1.Length == 1)
                                {
                                    Item_Value2 = Item_Value1[0].ToString().Split(',');   //check next value is different net
                                    Item_Value2 = Item_Value2[0].ToString().Split(' ');   //check next value is different net

                                    int aaa = table_net_new.Rows.Count;
                                    string ccc;
                                    for (int j = 0; j < Item_Value2.Length; j++)
                                    {
                                        if (Item_Value2[j].Trim() != "")
                                        {
                                            ccc = table_net_new.Rows[aaa - 1][1].ToString() + " " + Item_Value2[j].Trim();
                                            table_net_new.Rows[aaa - 1][1] = ccc;
                                        }
                                    }
                                }
                                else
                                {
                                    Item_Value2 = row[1].ToString().Split(',');   //check next value is different net

                                    if (Item_Value2.Length == 2)
                                    {
                                        int aaa = table_net_new.Rows.Count;

                                        row[1] = Item_Value2[0].Trim();
                                    }
                                    table_net_new.Rows.Add(row);
                                }
                            }

                            if ((temp_value != "$PACKAGES") && package_net_bit == 0)
                            {
                                Item_Value1 = lines[i].ToString().Split('!');
                                Item_Value2 = Item_Value1[1].ToString().Split(';');
                                string[] row = new string[Item_Value1.Length + Item_Value2.Length - 1];


                                row[0] = Item_Value1[0].Trim();

                                for (int j = 0; j < Item_Value2.Length; j++)
                                {
                                    row[j + 1] = Item_Value2[j].Trim();
                                }
                                table_new.Rows.Add(row);
                            }
                            break;
                    }
                }
            }
            else
                MessageBox.Show("InputFileError");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((textBox_File_Name1.Text == "") && (textBox_File_Name2.Text == ""))
                MessageBox.Show("Please set net file path");

            else
            {
                //table.Rows[1][1]

                //dataGridView1.Rows[0].Cells[1].Style.BackColor = Color.Red; //change color

                string Old_Location = dataGridView1.Rows[0].Cells[2].Value.ToString();
                string New_Location = dataGridView3.Rows[0].Cells[2].Value.ToString();
                string Old_Net;
                string New_Net;

                //Package Compare From Old to New
                for (int i = 0; i < dataGridView1.Rows.Count - 2; i++)
                {
                    Thread.Sleep(100);  //For ADLINK Office version
                    Old_Location = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView3.Rows.Count - 2; j++)
                    {
                        New_Location = dataGridView3.Rows[j].Cells[2].Value.ToString();
                        if (Old_Location == New_Location)
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 3; j++)
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
                    if (New_Row == -1)  //No find
                    {
                        for (int j = 0; j < 3; j++)
                            dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                    }
                    else
                    {
                        // Package check
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString() != dataGridView3.Rows[New_Row].Cells[0].Value.ToString())
                            dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;

                        // Part Value check
                        if (dataGridView1.Rows[i].Cells[1].Value.ToString() != dataGridView3.Rows[New_Row].Cells[1].Value.ToString())
                            dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Red;
                        //dataGridView1.Rows[i].Cells[1].Style.ForeColor

                        //Characters[start_pos, len].Font.Color;
                    }
                }

                //Net Compare From Old to New
                for (int i = 0; i < dataGridView2.Rows.Count - 2; i++)
                {
                    Old_Net = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView4.Rows.Count - 2; j++)
                    {
                        New_Net = dataGridView4.Rows[j].Cells[0].Value.ToString();
                        if (Old_Net == New_Net)
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 2; j++)
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.White;
                    if (New_Row == -1)  //No find
                    {
                        for (int j = 0; j < 2; j++)
                            dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                    }
                    else
                    {
                        // Net Name / Net List check
                        for (int j = 0; j < 2; j++)
                            if (dataGridView2.Rows[i].Cells[j].Value.ToString() != dataGridView4.Rows[New_Row].Cells[j].Value.ToString())
                                dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        //dataGridView2.Rows[i].Cells[j].Style.ForeColor = Color.Red;


                    }
                }

                //Package Compare From New to Old
                for (int i = 0; i < dataGridView3.Rows.Count - 2; i++)
                {
                    Old_Location = dataGridView3.Rows[i].Cells[2].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView1.Rows.Count - 2; j++)
                    {
                        New_Location = dataGridView1.Rows[j].Cells[2].Value.ToString();
                        if (Old_Location == New_Location)
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 3; j++)
                        dataGridView3.Rows[i].Cells[j].Style.BackColor = Color.White;
                    if (New_Row == -1)  //No find
                    {
                        for (int j = 0; j < 3; j++)
                            dataGridView3.Rows[i].Cells[j].Style.BackColor = Color.Green;
                    }
                    else
                    {
                        // Package check
                        if (dataGridView3.Rows[i].Cells[0].Value.ToString() != dataGridView1.Rows[New_Row].Cells[0].Value.ToString())
                            dataGridView3.Rows[i].Cells[0].Style.BackColor = Color.Red;

                        // Part Value check
                        if (dataGridView3.Rows[i].Cells[1].Value.ToString() != dataGridView1.Rows[New_Row].Cells[1].Value.ToString())
                            dataGridView3.Rows[i].Cells[1].Style.BackColor = Color.Red;
                    }
                }

                //Net Compare From New to Old
                for (int i = 0; i < dataGridView4.Rows.Count - 2; i++)
                {
                    Old_Net = dataGridView4.Rows[i].Cells[0].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView2.Rows.Count - 2; j++)
                    {
                        New_Net = dataGridView2.Rows[j].Cells[0].Value.ToString();
                        if (Old_Net == New_Net)
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 2; j++)
                        dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.White;
                    if (New_Row == -1)  //No find
                    {
                        for (int j = 0; j < 2; j++)
                            dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.Green;
                    }
                    else
                    {
                        // Net Name / Net List check
                        for (int j = 0; j < 2; j++)
                            if (dataGridView4.Rows[i].Cells[j].Value.ToString() != dataGridView2.Rows[New_Row].Cells[j].Value.ToString())
                                dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        //dataGridView4.Rows[i].Cells[j].Style.ForeColor = Color.Red;


                    }
                }

                //List<DataGridViewRow> resultRows = PerformSearch(dataGridView1, SearchList);
                MessageBox.Show("完成");
                button4.Enabled=true;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 儲存的Excel檔案路徑與檔案名稱
            const string PASSWORD = "12345";      // 設置的密碼
            string filePath = System.IO.Directory.GetCurrentDirectory() + "/" + "第一個Excel檔案";
            MessageBox.Show("檔案=" + filePath);

            //宣告名稱
            Excel.Workbook wBook;
            Excel.Worksheet wSheet;
            Excel.Range wRange;
            Excel.Application excelApp;
            excelApp = new Excel.Application();

            // 嘗試打開已經存在的workbook
            try
            {
                excelApp.Application.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, PASSWORD);
                //tbInfo.Text = tbInfo.Text + "Excel檔案已存在，讀入Excel檔案!\r\n";
            }
            catch (Exception ex)    //若檔案不存在則加入新的workbook
            {
                excelApp.Workbooks.Add(Type.Missing);
                //tbInfo.Text = tbInfo.Text + "新建立Excel檔案!\r\n";
            }

            /*****設定Excel檔案的屬性*****/
            // 讓Excel文件不可見 (不會顯示Application, 在背景工作)
            excelApp.Visible = false;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 取用第一個workbook
            wBook = excelApp.Workbooks[1];

            // 設定活頁簿焦點
            wBook.Activate();

            // 設定密碼
            //wBook.Password = PASSWORD;

            try
            {
                int sheetNum = wBook.Worksheets.Count;
                //tbInfo.Text = tbInfo.Text + string.Format(" 第{0}個Sheet\r\n", sheetNum);
                // 新增worksheet
                wSheet = (Excel.Worksheet)wBook.Worksheets.Add();

                // 設定worksheet的名稱
                wSheet.Name = string.Format("Old {0}", sheetNum);

                // 設定工作表焦點
                wSheet.Activate();

                // 設定第1列資料 (從1開始，不是從0)
                excelApp.Cells[1, 1] = "Old Pacakge";
                excelApp.Cells[1, 2] = "Old Part Value";
                excelApp.Cells[1, 3] = "Old Location";
                excelApp.Cells[1, 4] = "New Pacakge";
                excelApp.Cells[1, 5] = "New Part Value";
                excelApp.Cells[1, 6] = "New Location";


                excelApp.Cells[1, 8] = "Old Net Name";
                excelApp.Cells[1, 9] = "Old Netlist";
                excelApp.Cells[1, 10] = "New Net Name";
                excelApp.Cells[1, 11] = "New Netlist";

                // 設定第Cell[1, 1]至Cell[1,2]顏色 (兩個Cell間形成的矩形都會被設置)
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 11]];
                wRange.Select();
                wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
                wRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.DimGray);

                //同單元格多種字元顏色
                int start_pos = 0;
                int len = 2;
                wSheet.Cells[1, 1].Characters[start_pos, len].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                wSheet.Cells[1, 1].Characters[start_pos + len + 2, len].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                               
                // 自動調整欄寬
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 3]];
                wRange.Select();
                wRange.Columns.AutoFit();

                //轉出Package Grid資料
                for (int i = 0; i < dataGridView1.Rows.Count - 2; i++)
                {
                    for (int j = 0; j < 3; j++) //Output Parameter
                    {
                        excelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1].Interior.Color = ColorTranslator.ToOle(dataGridView1.Rows[i].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                     }

                    // Find New location information
                    for (int k = 0; k < dataGridView3.Rows.Count - 2; k++)
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString() == dataGridView3.Rows[k].Cells[2].Value.ToString())
                        {
                            for (int j = 0; j < 3; j++) //Output Parameter
                            {
                                excelApp.Cells[i + 2, j + 1 + 3] = dataGridView3.Rows[k].Cells[j].Value.ToString();
                                excelApp.Cells[i + 2, j + 1 + 3].Interior.Color = ColorTranslator.ToOle(dataGridView3.Rows[k].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                            }
                            break;
                        }
                    }
                }
                int Data_Count = dataGridView1.Rows.Count;
                for (int i = 0; i < dataGridView3.Rows.Count - 2; i++)  //Add New item when old non
                {
                    for (int k = 0; k < dataGridView1.Rows.Count - 2; k++)  //Add New item when old non
                    {
                        if (dataGridView3.Rows[i].Cells[2].Value.ToString() == dataGridView1.Rows[k].Cells[2].Value.ToString()) //if same no add new
                            break;
                        if (k + 1 == (dataGridView1.Rows.Count - 2))//new
                        { 
                            for (int j = 0; j < 3; j++) //Output Parameter
                            {
                                excelApp.Cells[Data_Count, j + 1 + 3] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                                excelApp.Cells[Data_Count, j + 1 + 3].Interior.Color = ColorTranslator.ToOle(dataGridView3.Rows[i].Cells[j].Style.BackColor);    
                            }
                            Data_Count++;
                        }
                    }
                }




                //轉出Net Grid資料
                for (int i = 0; i < dataGridView2.Rows.Count - 2; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        excelApp.Cells[i + 2, j + 1 + 7] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1 + 7].Interior.Color = ColorTranslator.ToOle(dataGridView2.Rows[i].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                        excelApp.Cells[i + 2, j + 1 + 9] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(dataGridView2.Rows[i].Cells[j].Style.BackColor);



                        if (dataGridView2.Rows[i].Cells[j].Style.BackColor == Color.Yellow) //if delete value is null
                        {
                            excelApp.Cells[i + 2, j + 1 + 9] = "";
                            excelApp.Cells[i + 2, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(Color.White);
                        }
                        if ((j == 1) && (dataGridView2.Rows[i].Cells[j].Style.BackColor == Color.Red))
                        {
                            string Old_Net = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            string Old_Net_list = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            string New_Net_list;
                            string New_Net;

                            dataGridView4.Rows[i].Cells[1].Value.ToString();


                            for (int k = 0; k < dataGridView4.Rows.Count - 2; k++)  //find New_Net_Row
                            {
                                New_Net = dataGridView4.Rows[k].Cells[0].Value.ToString();
                                if (Old_Net == New_Net)
                                {
                                    New_Net_list = dataGridView4.Rows[k].Cells[1].Value.ToString();
                                    wSheet.Cells[i + 2, j + 1 + 9] = dataGridView4.Rows[k].Cells[1].Value.ToString();

                                    string[] Item_Value1;
                                    string[] Item_Value2;
                                    Item_Value1 = Old_Net_list.ToString().Split(' ');
                                    Item_Value2 = New_Net_list.ToString().Split(' ');
                                    string New_Temp_Value = "";
                                    int start_pos_array = 1;

                                    
                                    for (int Old_Array_i = 0; Old_Array_i < Item_Value1.Length; Old_Array_i++)    //compare New array
                                    {
                                        //string Old_Temp_Value;
                                        int New_Row = -1; //check find bit is done
                                        New_Temp_Value = Item_Value1[Old_Array_i].Trim();
                                        
                                        for (int New_Array_i = 0; New_Array_i < Item_Value2.Length; New_Array_i++)  //find New array location
                                        {
                                            if (Item_Value2[New_Array_i].Trim() == New_Temp_Value)
                                            {
                                                New_Row++;
                                                break;
                                            }
                                        }
                                        //excelApp.Cells[i + 2, j + 1 + 4] = New_Temp_Value;
                                        if (New_Row < 0)
                                        {
                                            wSheet.Cells[i + 2, j + 1 + 7].Characters[start_pos_array, New_Temp_Value.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                                            wSheet.Cells[i + 2, j + 1 + 7].Characters[start_pos_array, New_Temp_Value.Length].Font.FontStyle = "bold";
                                         }
                                        start_pos_array = start_pos_array + New_Temp_Value.Length + 1;
                                    }

                                    start_pos_array = 1;
                                    for (int New_Array_i = 0; New_Array_i < Item_Value2.Length; New_Array_i++)    //compare old array
                                    {
                                        //string Old_Temp_Value;
                                        int New_Row = -1; //check find bit is done
                                        New_Temp_Value = Item_Value2[New_Array_i].Trim();

                                        for (int Old_Array_i = 0; Old_Array_i < Item_Value1.Length; Old_Array_i++)  //find New array location
                                        {
                                            if (Item_Value1[Old_Array_i].Trim() == New_Temp_Value)
                                            {
                                                New_Row++;
                                                break;
                                            }
                                        }
                                        //excelApp.Cells[i + 2, j + 1 + 4] = New_Temp_Value;
                                        if (New_Row < 0)
                                        {
                                            wSheet.Cells[i + 2, j + 1 + 9].Characters[start_pos_array, New_Temp_Value.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                                            wSheet.Cells[i + 2, j + 1 + 9].Characters[start_pos_array, New_Temp_Value.Length].Font.FontStyle = "bold";
                                        }
                                        start_pos_array = start_pos_array + New_Temp_Value.Length + 1;
                                    }

                                    break;
                                }
                            }



                        }

                    }


                }
                Data_Count = dataGridView2.Rows.Count;
                for (int i = 0; i < dataGridView4.Rows.Count - 2; i++)  //Add New item when old non
                {
                    for (int k = 0; k < dataGridView2.Rows.Count - 2; k++)  //Add New item when old non
                    {
                        if (dataGridView4.Rows[i].Cells[0].Value.ToString() == dataGridView2.Rows[k].Cells[0].Value.ToString()) //if same no add new
                            break;
                        if (k + 1 == (dataGridView2.Rows.Count - 2))//new
                        {
                            for (int j = 0; j < 2; j++) //Output Parameter
                            {
                                excelApp.Cells[Data_Count, j + 1 + 9] = dataGridView4.Rows[i].Cells[j].Value.ToString();
                                excelApp.Cells[Data_Count, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(dataGridView4.Rows[i].Cells[j].Style.BackColor);

                            }
                            Data_Count++;
                        }
                    }
                }





                try
                {
                    // 儲存workbook
                    wBook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //tbInfo.Text = tbInfo.Text + "成功儲存!\r\n";
                }
                catch (Exception ex)
                {
                    //tbInfo.Text = tbInfo.Text + "儲存失敗，請關閉該Excel檔案\r\n";
                }
            }
            catch (Exception ex)
            {
                //tbInfo.Text = tbInfo.Text + "生成時產生錯誤!\r\n";
            }

            //關閉workbook
            wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            MessageBox.Show("完成");
        }

        private void informationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" 作者:Sky \n 版本: 0.0 初版, 下次修改程式運行速度\n 修改日期: 2022/08/07");
        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" 準備中");
        }
    }
}
