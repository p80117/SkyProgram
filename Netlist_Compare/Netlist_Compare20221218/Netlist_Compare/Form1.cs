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

//確認檔案版本
using System.Diagnostics;
using System.Reflection;//記得using
using System.Security;

namespace Netlist_Compare
{
    public partial class Form1 : Form
    {
        public bool adlink { get; set; }
        
        public Form1()
        {
            InitializeComponent();
        }

        DataTable table = new DataTable();
        DataTable table_net = new DataTable();
        DataTable table_new = new DataTable();
        DataTable table_net_new = new DataTable();
        public void Form1_Load(object sender, EventArgs e)
        {
            adlink = false; //adlink 降速用
            if (Control.ModifierKeys != Keys.Shift)
            {
                Check_File_Version();   //檢查版本
                adlink = true; //adlink 降速用
            }


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

            dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView2.Columns[1].Width = 200;

            dataGridView4.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView4.Columns[1].Width = 200;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConstString_Path = "";
			int err=0;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ConstString_Path = dialog.FileName;
                textBox_File_Name1.Text = ConstString_Path;
                //MessageBox.Show(ConstString_Path);

            }
		
			DataTable table_Package = table;
			DataTable table_Net = table_net;
			
			err = Net_Import(ConstString_Path, table_Package, table_Net);	//2022.09.04
			if(err<0)MessageBox.Show("InputFileError");
            //CopyDataGridView(dataGridView1, dataGridView5); //2022.08.21.1

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string ConstString_Path = "";
			int err=0;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ConstString_Path = dialog.FileName;
                textBox_File_Name2.Text = ConstString_Path;
                //MessageBox.Show(ConstString_Path);

            }
            
			DataTable table_Package = table_new;
			DataTable table_Net = table_net_new;
			
			err = Net_Import(ConstString_Path, table_Package, table_Net);	//2022.09.04
			if(err<0)MessageBox.Show("InputFileError");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((textBox_File_Name1.Text == "") && (textBox_File_Name2.Text == ""))
                MessageBox.Show("Please set net file path");

            else
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;

                //table.Rows[1][1]

                //dataGridView1.Rows[0].Cells[1].Style.BackColor = Color.Red; //change color

                string Old_Location = dataGridView1.Rows[0].Cells[2].Value.ToString();
                string New_Location = dataGridView3.Rows[0].Cells[2].Value.ToString();
                string Old_Net;
                string New_Net;

                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView3, dataGridView5); //2022.08.21.1                

                //Package Compare From Old to New
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if(adlink)
                        Thread.Sleep(10);  //For ADLINK Office version
                    Old_Location = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    int New_Row = -1;


                    for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)  //Get New Data
                    {
                        New_Location = dataGridView5.Rows[j].Cells[2].Value.ToString();
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
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString() != dataGridView5.Rows[New_Row].Cells[0].Value.ToString())
                            dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;

                        // Part Value check
                        if (dataGridView1.Rows[i].Cells[1].Value.ToString() != dataGridView5.Rows[New_Row].Cells[1].Value.ToString())
                            dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Red;
                        //dataGridView1.Rows[i].Cells[1].Style.ForeColor

                        dataGridView5.Rows.Remove(dataGridView5.Rows[New_Row]);   //Delete compare data
                        //Characters[start_pos, len].Font.Color;
                    }
                }
                
                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView4, dataGridView5); //2022.08.21.1       
                
                //Net Compare From Old to New
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    Old_Net = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)  //find New net is same as Old
                    {
                        New_Net = dataGridView5.Rows[j].Cells[0].Value.ToString();
                        if (Old_Net == New_Net) //If New equare old, then exit search for
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 2; j++) //Default back color is white
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.White;
                    bool Red_Different = true;
                    if (New_Row == -1)  //No find on new back color is yellow
                    {
                        for (int j = 0; j < 2; j++)
                            dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                    }
                    else //Find New net is same as old, next to check value
                    {
                        // Net List check
                        {
                            string Old_Net_list = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            string New_Net_list = dataGridView5.Rows[New_Row].Cells[1].Value.ToString();

                            string[] Item_Value1;
                            string[] Item_Value2;
                            Item_Value1 = Old_Net_list.ToString().Split(' ');
                            Item_Value2 = New_Net_list.ToString().Split(' ');
                            string New_Temp_Value = "";                            

                            if (Item_Value1.Length == Item_Value2.Length) //如果【接點數量】不同，表示有差異，可以不用再確認
                                for (int Old_Array_i = 0; Old_Array_i < Item_Value1.Length; Old_Array_i++)    //compare New array
                                {
                                    //string Old_Temp_Value;
                                    New_Temp_Value = Item_Value1[Old_Array_i].Trim();
                                    Red_Different = true;
                                    for (int New_Array_i = 0; New_Array_i < Item_Value2.Length; New_Array_i++)  //find New array location
                                    {
                                        
                                        if (Item_Value2[New_Array_i].Trim() == New_Temp_Value)  //find same【接點】, then exit search for
                                        {
                                            Red_Different = false;
                                            break;
                                        }
                                        if ((New_Array_i == (Item_Value2.Length - 1)) && (Item_Value2[Item_Value2.Length - 1].Trim() != New_Temp_Value))
                                        {                                    
                                            /*如果沒有找到就要為紅色
                                             2.) 【Old接點】於【New接點】的值都沒有找到
                                             */
                                            Red_Different = true;
                                        }
                                    }
                                    //如果沒有找到為紅色就停止比對
                                    if (Red_Different == true)
                                        break;
                                }

                        }
                        if(Red_Different)
                            for (int j = 0; j < 2; j++)
                            //if (dataGridView2.Rows[i].Cells[j].Value.ToString() != dataGridView5.Rows[New_Row].Cells[j].Value.ToString())                         
                                dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        //dataGridView2.Rows[i].Cells[j].Style.ForeColor = Color.Red;

                        dataGridView5.Rows.Remove(dataGridView5.Rows[New_Row]);   //Delete compare data
                     }
                }

                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView1, dataGridView5); //2022.08.21.1     
                
                //Package Compare From New to Old
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    Old_Location = dataGridView3.Rows[i].Cells[2].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)
                    {
                        New_Location = dataGridView5.Rows[j].Cells[2].Value.ToString();
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
                        if (dataGridView3.Rows[i].Cells[0].Value.ToString() != dataGridView5.Rows[New_Row].Cells[0].Value.ToString())
                            dataGridView3.Rows[i].Cells[0].Style.BackColor = Color.Red;

                        // Part Value check
                        if (dataGridView3.Rows[i].Cells[1].Value.ToString() != dataGridView5.Rows[New_Row].Cells[1].Value.ToString())
                            dataGridView3.Rows[i].Cells[1].Style.BackColor = Color.Red;
                        dataGridView5.Rows.Remove(dataGridView5.Rows[New_Row]);   //Delete compare data
                    }
                }

                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView2, dataGridView5); //2022.08.21.1  
                
                //Net Compare From New to Old
                for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
                {
                    Old_Net = dataGridView4.Rows[i].Cells[0].Value.ToString();
                    int New_Row = -1;
                    for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)
                    {
                        New_Net = dataGridView5.Rows[j].Cells[0].Value.ToString();
                        if (Old_Net == New_Net)
                        {
                            New_Row = j;
                            break;
                        }
                    }
                    for (int j = 0; j < 2; j++)
                        dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.White;
                    bool Red_Different = true;
                    if (New_Row == -1)  //No find
                    {
                        for (int j = 0; j < 2; j++)
                            dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.Green;
                    }
                    else
                    {

                        // Net List check
                        {
                            string Old_Net_list = dataGridView4.Rows[i].Cells[1].Value.ToString();
                            string New_Net_list = dataGridView5.Rows[New_Row].Cells[1].Value.ToString();

                            string[] Item_Value1;
                            string[] Item_Value2;
                            Item_Value1 = Old_Net_list.ToString().Split(' ');
                            Item_Value2 = New_Net_list.ToString().Split(' ');
                            string New_Temp_Value = "";

                            if (Item_Value1.Length == Item_Value2.Length) //如果【接點數量】不同，表示有差異，可以不用再確認
                                for (int Old_Array_i = 0; Old_Array_i < Item_Value1.Length; Old_Array_i++)    //compare New array
                                {
                                    //string Old_Temp_Value;
                                    New_Temp_Value = Item_Value1[Old_Array_i].Trim();
                                    Red_Different = true;
                                    for (int New_Array_i = 0; New_Array_i < Item_Value2.Length; New_Array_i++)  //find New array location
                                    {

                                        if (Item_Value2[New_Array_i].Trim() == New_Temp_Value)  //find same【接點】, then exit search for
                                        {
                                            Red_Different = false;
                                            break;
                                        }
                                        if ((New_Array_i == (Item_Value2.Length - 1)) && (Item_Value2[Item_Value2.Length - 1].Trim() != New_Temp_Value))
                                        {
                                            /*如果沒有找到就要為紅色
                                             2.) 【Old接點】於【New接點】的值都沒有找到
                                             */
                                            Red_Different = true;
                                        }
                                    }
                                    //如果沒有找到為紅色就停止比對
                                    if (Red_Different == true)
                                        break;
                                }

                        }
                        if (Red_Different)
                            // Net Name / Net List check
                            for (int j = 0; j < 2; j++)
                            //if (dataGridView4.Rows[i].Cells[j].Value.ToString() != dataGridView5.Rows[New_Row].Cells[j].Value.ToString())
                                dataGridView4.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        dataGridView5.Rows.Remove(dataGridView5.Rows[New_Row]);   //Delete compare data
                        //dataGridView4.Rows[i].Cells[j].Style.ForeColor = Color.Red;


                    }
                }

                //New Package new only different item
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (dataGridView3.Rows[i].Cells[1].Style.BackColor == Color.White)
                    {
                        dataGridView3.Rows.Remove(dataGridView3.Rows[i]);   //Delete compare data
                        i = i - 1;
                    }

                }

                //New Net only different item
                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    if(dataGridView4.Rows[i].Cells[1].Style.BackColor == Color.White)
                    {
                        dataGridView4.Rows.Remove(dataGridView4.Rows[i]);   //Delete compare data
                        i = i - 1;
                    }

                }
				
				//Compare only netname different item
				for (int i = 0; i< dataGridView2.Rows.Count-1; i++)
				{					
					if(dataGridView2.Rows[i].Cells[1].Style.BackColor == Color.Yellow)
					{
						string Old_Net_Value = dataGridView2.Rows[i].Cells[1].Value.ToString();
						for(int j = 0; j< dataGridView4.Rows.Count-1; j++)
						{
							string New_Net_Value = dataGridView4.Rows[j].Cells[1].Value.ToString();
							if(Old_Net_Value == New_Net_Value)
							{
								for(int k=0; k<2; k++)
								{
									dataGridView2.Rows[i].Cells[k].Style.BackColor = Color.Orange;
									dataGridView4.Rows[j].Cells[k].Style.BackColor = Color.Orange;
								}								
								break;
							}
							
						}
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
                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView3, dataGridView5); //2022.08.21.1                 
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < 3; j++) //Output Parameter
                    {
                        excelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1].Interior.Color = ColorTranslator.ToOle(dataGridView1.Rows[i].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                     }

                    // Find New location information
                    for (int k = 0; k < dataGridView5.Rows.Count - 1; k++)
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString() == dataGridView5.Rows[k].Cells[2].Value.ToString())
                        {
                            for (int j = 0; j < 3; j++) //Output Parameter
                            {
                                excelApp.Cells[i + 2, j + 1 + 3] = dataGridView5.Rows[k].Cells[j].Value.ToString();
                                excelApp.Cells[i + 2, j + 1 + 3].Interior.Color = ColorTranslator.ToOle(dataGridView5.Rows[k].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                            }
                            dataGridView5.Rows.Remove(dataGridView5.Rows[k]);   //Delete compare data
                            break;
                        }
                    }
                    
                }
                int Data_Count = dataGridView1.Rows.Count;
                for (int i = 0; i < dataGridView5.Rows.Count - 1; i++)  //Add New item when old non
                {
                    //for (int k = 0; k < dataGridView1.Rows.Count - 2; k++)  //Find old item is same new item
                    {
                        //if (dataGridView5.Rows[i].Cells[2].Value.ToString() == dataGridView1.Rows[k].Cells[2].Value.ToString()) //if same no add new
                        //    break;  //已無需要此判斷
                        //if (k + 1 == (dataGridView1.Rows.Count - 2))//new
                        { 
                            for (int j = 0; j < 3; j++) //Output Parameter
                            {
                                excelApp.Cells[Data_Count, j + 1 + 3] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                                excelApp.Cells[Data_Count, j + 1 + 3].Interior.Color = ColorTranslator.ToOle(dataGridView5.Rows[i].Cells[j].Style.BackColor);    
                            }
                            Data_Count++;
                        }
                    }
                }
                               
                //轉出Net Grid資料
                dataGridView5.Columns.Clear();  //Clear all data
                CopyDataGridView(dataGridView4, dataGridView5); //2022.08.21.1     
                
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        excelApp.Cells[i + 2, j + 1 + 7] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1 + 7].Interior.Color = ColorTranslator.ToOle(dataGridView2.Rows[i].Cells[j].Style.BackColor);//System.Drawing.Color.DimGray);
                        excelApp.Cells[i + 2, j + 1 + 9] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        excelApp.Cells[i + 2, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(dataGridView2.Rows[i].Cells[j].Style.BackColor);
					}
					if (dataGridView2.Rows[i].Cells[1].Style.BackColor == Color.Yellow) //if delete value is null
					{
						excelApp.Cells[i + 2, 0 + 1 + 9] = "";
						excelApp.Cells[i + 2, 1 + 1 + 9] = "";
						//excelApp.Cells[i + 2, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(Color.White);
					}
					if (dataGridView2.Rows[i].Cells[1].Style.BackColor == Color.Red)
					{
						string Old_Net = dataGridView2.Rows[i].Cells[0].Value.ToString();
						string Old_Net_list = dataGridView2.Rows[i].Cells[1].Value.ToString();
						string New_Net_list;
						string New_Net;

						//dataGridView4.Rows[i].Cells[1].Value.ToString();


						for (int k = 0; k < dataGridView5.Rows.Count - 1; k++)  //find New_Net_Row
						{
							New_Net = dataGridView5.Rows[k].Cells[0].Value.ToString();
							if (Old_Net == New_Net)
							{
								New_Net_list = dataGridView5.Rows[k].Cells[1].Value.ToString();
								wSheet.Cells[i + 2, 1 + 1 + 9] = dataGridView5.Rows[k].Cells[1].Value.ToString();

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
										wSheet.Cells[i + 2, 1 + 1 + 7].Characters[start_pos_array, New_Temp_Value.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
										wSheet.Cells[i + 2, 1 + 1 + 7].Characters[start_pos_array, New_Temp_Value.Length].Font.FontStyle = "bold";
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
										wSheet.Cells[i + 2, 1 + 1 + 9].Characters[start_pos_array, New_Temp_Value.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
										wSheet.Cells[i + 2, 1 + 1 + 9].Characters[start_pos_array, New_Temp_Value.Length].Font.FontStyle = "bold";
									}
									start_pos_array = start_pos_array + New_Temp_Value.Length + 1;
								}

								dataGridView5.Rows.Remove(dataGridView5.Rows[k]);   //Delete compare data
								break;
							}
						}



					}
					if (dataGridView2.Rows[i].Cells[1].Style.BackColor == Color.Orange) //Find Different netname	//2022.09.04
					{
						string old_net_value = dataGridView2.Rows[i].Cells[1].Value.ToString();
						
						for (int k = 0; k < dataGridView5.Rows.Count - 1; k++)  //find New_Net_Row
						{
							string new_net_value = dataGridView5.Rows[k].Cells[1].Value.ToString();
							if( old_net_value == new_net_value)
							{
								excelApp.Cells[i + 2, 0 + 1 + 9] = dataGridView5.Rows[k].Cells[0].Value.ToString();	//update new net name
								dataGridView5.Rows.Remove(dataGridView5.Rows[k]);   //Delete compare data
								break;
							}
						}
					}
                    


                }
                Data_Count = dataGridView2.Rows.Count+1;
                for (int i = 0; i < dataGridView5.Rows.Count - 1; i++)  //Add New item when old non
                {
                    //for (int k = 0; k < dataGridView2.Rows.Count - 2; k++)  //Add New item when old non
                    {
                        //if (dataGridView4.Rows[i].Cells[0].Value.ToString() == dataGridView2.Rows[k].Cells[0].Value.ToString()) //if same no add new
                        //    break;
                        //if (k + 1 == (dataGridView2.Rows.Count - 2))//new
                        {
                            for (int j = 0; j < 2; j++) //Output Parameter
                            {
                                excelApp.Cells[Data_Count, j + 1 + 9] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                                excelApp.Cells[Data_Count, j + 1 + 9].Interior.Color = ColorTranslator.ToOle(dataGridView5.Rows[i].Cells[j].Style.BackColor);

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
            MessageBox.Show(" 作者:Sky " +
                "\n 版本: 2022.08.07.1 初版" +
                "\n 版本: 2022.08.20.1 版本控制" +
                "\n 版本: 2022.08.20.2 比過的不重複比較" +
                "\n 版本: 2022.08.21.1" +
                "\n        1.) New net, package 只保留有差異的部分" +
                "\n        2.) Compare only run 1 times" +
                "\n        3.) 修正Net排序不同，判斷為【差異】結果問題" +
                "\n 版本: 2022.09.04.1" +
                "\n        1.) 排除Net排序不同判斷差異" +
                "\n        2.) 新增Sky_Mode" +
                "\n        3.) 新增只有[Net Name]差異 => [橘色]" +

                "\n\n 待更新版本: 2022.09.??.1 " +
                "\n1.) 下次修改程式運行速度" +
                "\n2.) 進階比較，Net Name不同，netlist只有些微不同" +
                "\n2.) 軟體自動更新", "版本說明");

        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" 準備中");
        }

        private void checkVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Check_File_Version();
        }
        private void Check_File_Version()
        {

            //參考 : https://www.796t.com/content/1550451801.html
            //參考: http://www.aspphp.online/bianchen/dnet/cxiapu/cxprm/201701/191794.html

            //版本號設定: https://www.demo.tc/post/automatic%20versions%20%E5%88%A5%E5%86%8D%E6%89%8B%E5%8B%95%E6%94%B9%E7%89%88%E6%9C%AC%E8%99%9F%E4%BA%86
            /*
             1.) 於Properties中的AssemblyInfo.cs設定[assembly: AssemblyVersion("1.0.0.*")]

             */
            /*
                        string FileVersions = "123"; 
                        try
                        {
                            FileVersionInfo.GetVersionInfo(Path.Combine(Environment.CurrentDirectory, "Netlist_Compare.exe"));
                            //FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(Environment.CurrentDirectory + "\\Netlist_Compare.exe");//要獲取版本號的exe程式
                            FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);//取得此程式相關資訊

                            // Print the file name and version number. 
                            MessageBox.Show("File: " + myFileVersionInfo.FileDescription + '\n'+"Version: " + myFileVersionInfo.FileVersion);               
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Fail");
                            FileVersions = "";
                        }
             */
            string currentAssemblyPath = Environment.CurrentDirectory;
            string updatedAssemblyPath = @"\\adlink-fs1\BEC\HW5\@SEC_2\Sky\Tool";// C:\Users\sky\Desktop\TestProgram\Netlist_Compare2";

            try
            {             
                
            AssemblyName currentAssemblyName = AssemblyName.GetAssemblyName(currentAssemblyPath+"\\Netlist_Compare.exe");
            
            AssemblyName updatedAssemblyName = AssemblyName.GetAssemblyName(updatedAssemblyPath + "\\Netlist_Compare.exe");

            //Assembly currentAssembly = Assembly.LoadFile(currentAssemblyPath);
            //Assembly updatedAssembly = Assembly.LoadFile(updatedAssemblyPath);

            //AssemblyName currentAssemblyName = currentAssembly.GetName();
            //AssemblyName updatedAssemblyName = updatedAssembly.GetName();

            // 比較版本號
            if (updatedAssemblyName.Version.CompareTo(currentAssemblyName.Version) <= 0)
            {
                //MessageBox.Show("currentAssemblyName.Version : " + currentAssemblyName.Version + "\n" +
                //    "updatedAssemblyName.Version : " + updatedAssemblyName.Version
                //    );
                MessageBox.Show("版本為最新");
                // 不需要更新
                return;
            }
            // 更新
            MessageBox.Show("請更新版本");
                System.Diagnostics.Process.Start("Explorer.exe", updatedAssemblyPath);
            this.Close();
                //File.Copy(updatedAssemblyPath, currentAssemblyPath, true);  //無法直接複製, 需要排除

            }
            catch (FileNotFoundException)   //如果更新路徑沒有檔案，則無授權使用
            {
                /*
                 *例外狀況
//ArgumentNullException
assemblyFile 為 null。

//ArgumentException
assemblyFile 無效，如具有無效文化特性的組件。

FileNotFoundException
找不到 assemblyFile。

//SecurityException
呼叫端沒有路徑探索權限。

//BadImageFormatException
assemblyFile 不是有效的組件。

FileLoadException
已使用兩組不同的辨識項載入組件或模組兩次。
                 */

                MessageBox.Show("檔案無授權");
                this.Close();
            }
        }

        private DataGridView CopyDataGridView(DataGridView dgv_org, DataGridView dgv_copy)
        {
            //DataGridView dgv_copy = new DataGridView();
            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv_org.Columns)
                    {
                        dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                DataGridViewRow row = new DataGridViewRow();

                for (int i = 0; i < dgv_org.Rows.Count; i++)
                {
                    row = (DataGridViewRow)dgv_org.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dgv_org.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dgv_copy.Rows.Add(row);
                }
                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Copy DataGridViw: "+ex);
            }
            return dgv_copy;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //dataGridView5.DataSource = dataGridView1.DataSource;    //2022.08.21.1

            //dataGridView5.Rows.Remove(dataGridView5.Rows[1]);   //Delete Row
            //dataGridView5.DataSource = null;    //Clear Datagrid
            //dataGridView5.Rows.Clear();
            /*            dataGridView5.Columns.Clear();  //Clear all data

                        //CopyDataGridView(dataGridView1, dataGridView5); //2022.08.21.1


                        dataGridView5.Columns.Clear();  //Clear all data
                        CopyDataGridView(dataGridView1, dataGridView5); //2022.08.21.1   */
            if (Control.ModifierKeys == Keys.Shift)
            {
                MessageBox.Show("hi");
            }
        }
    
		private int Net_Import(string ConstString_Path, DataTable table_Package, DataTable table_Net)
		{
            // string ConstString_Path = textBox_File_Name1.Text;
			// DataGridView table_Package = table;
			// DataGridView table_Net = table_net;
			
            if (ConstString_Path != "")
            {   
                string[] lines = File.ReadAllLines(ConstString_Path);
                string[] Item_Value_Array1;
                string[] Item_Value_Array2;
				string[] Net_Array;
                int package_net_bit = 0;
                for (int i = 0; i < lines.Length; i++)
                {
                    string temp_value = lines[i].ToString();	//Get Origin netlist data to check [Package] or [Netlist]
                    switch (temp_value)
                    {
                        case "$NETS":
                        case "$END":
                            package_net_bit++;
                            break;

                        default:

                            if (package_net_bit == 1)	//Netlist
                            {
                                Item_Value_Array1 = lines[i].ToString().Split(';');

                                string[] row = new string[Item_Value_Array1.Length];

                                //row[0] = Item_Value_Array1[0].Trim();

                                for (int j = 0; j < Item_Value_Array1.Length; j++)
                                    row[j] = Item_Value_Array1[j].Trim();

                                if (Item_Value_Array1.Length == 1)
                                {
                                    Item_Value_Array2 = Item_Value_Array1[0].ToString().Split(',');   //check next value is different net
                                    Item_Value_Array2 = Item_Value_Array2[0].ToString().Split(' ');   //check next value is different net

                                    int aaa = table_Net.Rows.Count;
                                    string ccc;
                                    for (int j = 0; j < Item_Value_Array2.Length; j++)
                                    {
                                        if (Item_Value_Array2[j].Trim() != "")
                                        {
                                            ccc = table_Net.Rows[aaa - 1][1].ToString() + " " + Item_Value_Array2[j].Trim();
                                            //ccc = table_Net[aaa - 1, 1].Value + " " + Item_Value_Array2[j].Trim();
                                            table_Net.Rows[aaa - 1][1] = ccc;
                                            //table_Net[aaa - 1, 1].Value = ccc;
                                        }
                                    }
                                }
                                else
                                {	//new net								

                                    Item_Value_Array2 = row[1].ToString().Split(',');   //check next value is different net

                                    if (Item_Value_Array2.Length == 2)
                                    {
                                        int aaa = table_Net.Rows.Count;
                                        //int bbb = table_Net.Columns.Count;

                                        //string ccc = table_Net.Rows[aaa][1].ToString();

                                        row[1] = Item_Value_Array2[0].Trim();

                                    }
                                    table_Net.Rows.Add(row);
                                }
                            }

                            if ((temp_value != "$PACKAGES") && package_net_bit == 0)	//package
                            {
								Item_Value_Array1 = lines[i].ToString().Split('!');	//split [Package] [Value; Location]
                                Item_Value_Array2 = Item_Value_Array1[1].ToString().Split(';'); //split [Value] [Location]
                                string[] row = new string[Item_Value_Array1.Length + Item_Value_Array2.Length - 1];


                                row[0] = Item_Value_Array1[0].Trim();

                                for (int j = 0; j < Item_Value_Array2.Length; j++)
                                {
                                    row[j + 1] = Item_Value_Array2[j].Trim();
                                }
                                table_Package.Rows.Add(row);
                            }
                            break;
                    }
                }
				// sort string array						
				for(int Net_Row=0; Net_Row<table_Net.Rows.Count; Net_Row++)
				{
					string net_value;
                    net_value = table_Net.Rows[Net_Row][1].ToString();//net_value = (string)table_Net[Net_Row, 1].Value;
                    Net_Array = net_value.ToString().Split(' ');
					Array.Sort( Net_Array );
					net_value= Net_Array[0].Trim();
					for(int i=1; i<Net_Array.Length; i++)					
						net_value = net_value + " " + Net_Array[i].Trim();
					table_Net.Rows[Net_Row][1] = net_value;
                    //table_Net[Net_Row, 1].Value = net_value;

                }
                return 0;


            }
            else
                return -1;//MessageBox.Show("InputFileError");

		}
	}
}
