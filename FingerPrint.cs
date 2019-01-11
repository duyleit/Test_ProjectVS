using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Data.SqlClient;


namespace FingerPrint
{
    public partial class FingerPrint : Form
    {
        public Timer timer = new Timer();
        public string ServerHost = "";
        public OleDbConnection con = new OleDbConnection();
        public SqlConnection consql = new SqlConnection();

        public FingerPrint()
        {
            InitializeComponent();
           lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Run.......");
            timer.Tick += new EventHandler(timer_Tick); // Every time timer ticks, timer_Tick will be called
            timer.Interval = 1000;                      // milliseconds 
            timer.Start();                              // Start the timer
        }

        private void bt_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private OleDbDataAdapter getDataAdapter(String ASql)
        {
            OleDbCommand command = new OleDbCommand(ASql, con);
            OleDbDataAdapter Dad1 = new OleDbDataAdapter(command);
            return Dad1;
        }

        private DataSet GetTableTemp(string qry,string strcon)
        {
            con.Close();
            con = getcon(strcon);
            DataSet dsId = new DataSet();
            OleDbDataAdapter daId = new OleDbDataAdapter();
            daId = getDataAdapter(qry);
            daId.Fill(dsId);
            con.Close();
            return dsId;
        }
        private DataSet GetTableTempId(string qry, string init, string name)
        {
            con.Close();
            string strId = GetIniValue(name, "DATABASE", init).ToString().Trim();
            con = getcon(strId);
            DataSet dsId = new DataSet();
            OleDbDataAdapter daId = new OleDbDataAdapter();
            daId = getDataAdapter(qry);
            daId.Fill(dsId);
            con.Close();
            return dsId;
        }
        private void bt_run_Click(object sender, EventArgs e)
        {
            timer.Stop();
            Start(DateTime.Now);
        }

        public void Start( DateTime t)
        {
            string IniName = Application.StartupPath + @"\connect.ini";
            string[] arrDBname = { "STAR", "HAMS", "MITACOSQL" };
            string[] arrData = { };
            string name =t.ToString("yyyyMMdd");
           
            for (int ii = 0; ii < arrDBname.Length; ii++) // so DB tong can duyet
            {
                string strtbl = "";
                if (ii == 2)
                {
                    strtbl = GetIniValue(arrDBname[ii], "SERVER", IniName).ToString().Trim();
                }
                else
                {
                    strtbl = GetIniValue(arrDBname[ii], "DATABASE", IniName).ToString().Trim();
                }

                if (strtbl == "")
                {
                    continue;
                }
                arrData = strtbl.Split(',');
                //progressBar1.Value = 0;
                //progressBar1.Maximum = 0;
                //progressBar1.Step = 1;
                OleDbCommand command = new OleDbCommand();
                OleDbDataReader reader;
                StreamWriter myFile;
                string query = "";
              
                for (int i = 0; i < arrData.Length; i++) // so DB con can duyet
                {
                    if (ii == 2) // connet sql server
                    {
                        SqlConnection con = getconsql(arrData[i]);
                    }
                    else  // connect ms access
                    {
                        con = getcon(arrData[i]);
                    }
                   
                    // string query = "SELECT  [eventDate],[eventTime],' ' as [In],[eventCard],[deviceID],' ' as [Out] FROM " + arrData[i] + " WHERE eventData = '" + DateTime.Now.ToString("yyyy/MM/dd")+"'" ;
                    // string query = "SELECT  [eventDate] +' '+[eventTime]+',' as [DateTime],[eventCard],[deviceID]+ ',,' as [deviceID]  FROM PubEvent WHERE eventDate = '" + dtpicker_date.Text + "'";
                    if (ii == 0) // STAR DB
                    {
                     //   DataSet dsId = GetTableTempId("select * from [data$]", IniName, "ID");  // Table temp get ID
                      //  DataSet ds = GetTableTemp("select IDX, ORGPOLLINGDATA from T_ORGPOLLINGDATA where Len(ORGPOLLINGDATA)>=33 and Mid(ORGPOLLINGDATA,15,8)='" + dtpicker_date.Value.Date.ToString("yyyMMdd") + "'", IniName, "STAR", arrData[i]); // Table temp get Star
                        DataSet ds = GetTableTemp("select IDX, ORGPOLLINGDATA from T_ORGPOLLINGDATA where Len(ORGPOLLINGDATA)>=33 and Mid(ORGPOLLINGDATA,15,8)='" + name + "'", arrData[i]); // Table temp get Star
                        string info = "";
                        bool flginfo = false;
                        string dd = "";
                        string hh = "";
                        string id = "";
                        StringBuilder getdata = new StringBuilder();
                        for (int r = 0; r < ds.Tables[0].Rows.Count; r++)
                        {
                            dd = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 15, 4).ToString() + '/' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 19, 2).ToString() + '/' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 21, 2).ToString(); ;
                            hh = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 24, 2) + ':' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 26, 2) + ':' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 28, 2);
                            id = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 7, 8);
                            //for (int rid = 0; rid < dsId.Tables[0].Rows.Count; rid++)
                            //{
                            //    if (id == dsId.Tables[0].Rows[rid][2].ToString())
                            //    {
                            //        id = Microsoft.VisualBasic.Strings.Trim(dsId.Tables[0].Rows[rid][0].ToString());
                            //        break;
                            //    }
                            //}
                            id = Microsoft.VisualBasic.Strings.Trim(Scalar("select MANV  from [data$] where MSTHETU = '" + id + "'", getcon(GetIniValue("ID", "DATABASE", IniName).ToString().Trim())));  //Nang cap  30112018
                            info = dd + ' ' + hh + ",," + id + ",00,,";
                            flginfo = true;
                            getdata.Append(info + "\r\n");
                            //myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                            //myFile.WriteLine(String.Format("{0}", info));
                            //flginfo = true;
                            //myFile.Flush();
                            //myFile.Close();
                         //   progressBar1.PerformStep();
                        }
                        myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                        myFile.WriteLine(getdata.ToString());
                        myFile.Flush();
                        myFile.Close();
                        if (flginfo)
                        {
                            lbox_info.Items.Add(t.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + " !!!");
                        }
                        else
                        {
                            lbox_info.Items.Add(t.ToShortTimeString() + ", Not data " + arrData[i] + " !!!");
                        }
                    }
                    else if (ii == 1) //HAMS DB
                    {
                        //query = "SELECT  ORGPOLLINGDATA as [info]  FROM T_ORGPOLLINGDATA";// WHERE eventDate = '" + dtpicker_date.Text + "'";
                        //  query = "SELECT  Mid(ORGPOLLINGDATA,15,8)+ ' '+ Mid(ORGPOLLINGDATA,24,6)+', ,'+  Mid(ORGPOLLINGDATA,7,8)  + ',00, ,' as [info]FROM T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) LIKE '*"+dtpicker_date.Value.Date.ToString("yyyMMdd") +"*'";
                        //  query = "SELECT  Mid(ORGPOLLINGDATA,15,8)+ ' '+ Mid(ORGPOLLINGDATA,24,6)+', ,'+  Mid(ORGPOLLINGDATA,7,8)  + ',00, ,' as [info]FROM T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) LIKE '*20181101*'";
                        query = "SELECT  [eventDate] +' '+[eventTime]+',,'+[eventCard]+','+[deviceID]+ ',,' as [info]  FROM PubEvent WHERE eventDate = '" + t.ToString("yyyy/MM/dd") + "'";
                        //string connectionSql = "Server=(local);Database=AdventureWorks2016CTP3;Integrated Security=true";
                        // private static FileStream fs = new FileStream(@"d:\temp\mcb.txt", FileMode.OpenOrCreate, FileAccess.Write);  
                        // private static StreamWriter m_streamWriter = new StreamWriter(fs);
                        //   StreamWriter myFile = new StreamWriter(@"D:\\FingerPrint_Project\\FingerPrint\\"+name+".txt");
                        // con = getcon(arrData[i]); moi xoa 22112018
                        StringBuilder getdata = new StringBuilder();
                        using (con)
                        {
                            command.CommandText = query;
                            command.Connection = con;
                            reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {
                                try
                                {  
                                   while (reader.Read())
                                    {
                                        //myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                        //myFile.WriteLine(String.Format("{0}", reader["info"]));
                                        //myFile.Flush();
                                        //myFile.Close();
                                        //   progressBar1.PerformStep();
                                        getdata.Append(reader["info"] +"\r\n");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    Close();
                                }
                                finally
                                {

                                    myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                    myFile.WriteLine(getdata.ToString());
                                    myFile.Flush();
                                    myFile.Close();
                                    reader.Close();
                                    lbox_info.Items.Add(t.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + " !!!");
                                }
                            }
                            else
                            {
                                lbox_info.Items.Add(t.ToShortTimeString() + ", Not data " + arrData[i] + " !!!");
                            }
                        }
                    }
                    else  //MITACO DB
                    {
                        query = "SELECT  NgayCham,GioCham,MaChamCong,MaSoMay FROM CheckInOut WHERE NgayCham = '" + t.ToString("yyyy-MM-dd") + "'";
                        //string connectionSql = "Server=(local);Database=AdventureWorks2016CTP3;Integrated Security=true";
                        // private static FileStream fs = new FileStream(@"d:\temp\mcb.txt", FileMode.OpenOrCreate, FileAccess.Write);  
                        // private static StreamWriter m_streamWriter = new StreamWriter(fs);
                        //   StreamWriter myFile = new StreamWriter(@"D:\\FingerPrint_Project\\FingerPrint\\"+name+".txt");
                        // con = getcon(arrData[i]); moi xoa 22112018
                        StringBuilder getdata = new StringBuilder();
                        using (consql)
                        {
                            string info;
                            string dd = "";
                            string hh = "";
                            SqlCommand commandsql = new SqlCommand();
                            SqlDataReader readersql;
                            commandsql.CommandText = query;
                            commandsql.Connection = consql;
                            readersql = commandsql.ExecuteReader();
                            if (readersql.HasRows)
                            {
                                try
                                {
                                   while (readersql.Read())
                                    {
                                        dd = ((DateTime)readersql["NgayCham"]).ToString("yyyy/MM/dd");
                                        hh = ((DateTime)readersql["GioCham"]).ToString("HH:mm:ss");
                                        info = dd + ' ' + hh + ",," + readersql["MaChamCong"].ToString() + ',' + readersql["MaSoMay"].ToString() + ",,";
                                        //myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                        //myFile.WriteLine(String.Format("{0}", info));
                                        //myFile.Flush();
                                        //myFile.Close();
                                        getdata.Append(info + "\r\n");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    Close();
                                }
                                finally
                                {
                                    myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                    myFile.WriteLine(getdata.ToString());
                                    myFile.Flush();
                                    myFile.Close();
                                    readersql.Close();
                                    lbox_info.Items.Add(t.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + '\\' + arrDBname[2] + " !!!");
                                 }
                            }
                            else
                            {
                                lbox_info.Items.Add(t.ToShortTimeString() + ", Not data " + arrData[i] + '\\' + arrDBname[2] + " !!!");
                            }
                        }

                    }

                }
                con.Close();
            }
            Reset();
            timer.Start();
        }

        public void Reset()
        {
            string IniName = Application.StartupPath + @"\connect.ini";
            string strtbl = GetIniValue("TIME", "INIT", IniName).ToString().Trim();
            lbl_timer.Text = strtbl;
            lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Run.......");
         }

        public void timer_Tick(object sender, EventArgs e)
        {
            string IniName = Application.StartupPath + @"\connect.ini";
            string strtbl = GetIniValue("TIMESPAN", "INIT", IniName).ToString().Trim();

            lbl_timer.Text = ((Convert.ToInt16(lbl_timer.Text) - 1)).ToString();
        //   TimeSpan t1 = TimeSpan.Parse("00:00:00");
        //   TimeSpan t2 = TimeSpan.Parse(DateTime.Now.ToString("HH:mm:ss"));
        //   TimeSpan t2 = TimeSpan.Parse(DateTime.Now.ToString("00:45:00")); // test
        //   string minu =Convert.ToString( Convert.ToInt16(strtbl) * 2);
        //   TimeSpan t = t2 - t1;
            TimeSpan t = TimeSpan.Parse(DateTime.Now.ToString("HH:mm:ss"));
            if (Convert.ToInt16(lbl_timer.Text) == 0)
            {
                if (t < TimeSpan.Parse(strtbl))
                {
                    File.Delete(Application.StartupPath + "\\dataExport\\" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + ".txt");
                    Start(DateTime.Now.AddDays(-1));
                }
                File.Delete(Application.StartupPath + "\\dataExport\\" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                Start(DateTime.Now);
            }
        }

        public class IniFile
        {
            public string path;

            //[DllImport("kernel32")]
            //private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
            [DllImport("kernel32")]
            private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

            /// INIFile Constructor.
            public IniFile(string INIPath)
            {
                path = INIPath;
            }
            // Write Data to the INI File
            //public void IniWriteValue(string Section, string Key, string Value)
            //{
            //    WritePrivateProfileString(Section, Key, Value, this.path);
            //}
            // Read Data Value From the Ini File
            public string IniReadValue(string Section, string Key) // ham nay thuc hien xong se close reader
            {
                StringBuilder temp = new StringBuilder(3000);
                //string val = "";
                int i = GetPrivateProfileString(Section, Key, "", temp, 3000, this.path);
                //int yy = temp.Length;
                //  string  val = "\\" + temp.ToString();
                return temp.ToString();
                //return "\\192.168.0.100"   ;      
            }
        }

        private OleDbConnection getcon(string strpath)
        {
            //  OleDbConnection con;
            //  con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\" + ServerHost + "\\HAMS\\HAMS_2018.mdb");
            //  con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|HAMS_2018.mdb");

            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            string ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strpath;

            if (Microsoft.VisualBasic.Strings.Right(strpath, 4) == ".xls" || Microsoft.VisualBasic.Strings.Right(strpath, 5) == ".xlsx")
            {
                ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strpath;
                ConnStr += @";Extended Properties='Excel 8.0;HDR=Yes;'";
            }

            con.ConnectionString = ConnStr;

            try
            {
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
            return con;
        }
        private SqlConnection getconsql(string strpath)
        {
            string DBUser = "vinhtuyen";
            string DBPwd = "VinhTuyen.1";
            string DBName = "MITACOSQL";
            string ConnStr = "";
            if (consql.State == ConnectionState.Open)
            {
                consql.Close();
            }


            ConnStr = "Data Source = " + strpath + " ;Initial Catalog= " + DBName + " ;User ID = " + DBUser + " ;Password = " + DBPwd;


            consql.ConnectionString = ConnStr;

            try
            {
                if (consql.State == ConnectionState.Closed)
                {
                    consql.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
            return consql;
        }

        private string GetIniValue(String section, String keyname, string strIniFile)
        {
            IniFile ini = new IniFile(strIniFile);
            return ini.IniReadValue(section, keyname);
        }

        private bool setDBConn()
        {
            string IniName = Application.StartupPath + @"\connect.ini";
            ServerHost = GetIniValue("CONNECT", "SERVER", IniName);
            return true;
        }

        private void InitForm()
        {
            string IniName = Application.StartupPath + @"\connect.ini";
            string strtbl = GetIniValue("TIME", "INIT", IniName).ToString().Trim();
            lbl_timer.Text = strtbl;
            if (!setDBConn())
            {
                this.Close();
            }
        }

        private void FingerPrint_Load(object sender, EventArgs e)
        {
            InitForm();
            //getcon();
        }

        private string Scalarsql(string qry, SqlConnection con)
        {
            string result = "";
            SqlCommand command = new SqlCommand();
            command.CommandText = qry;
            command.Connection = con;
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            result = command.ExecuteScalar().ToString();
            con.Close();
            return result;
        }

        private string Scalar(string qry, OleDbConnection con)
        {
            string result = "";
            OleDbCommand command = new OleDbCommand();
            command.CommandText = qry;
            command.Connection = con;
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            result = (string)command.ExecuteScalar();
            con.Close();
            return result;
        }

        private OleDbDataReader ExcReader(string qry, OleDbConnection con)
        {
            OleDbDataReader reader;
            OleDbCommand command = new OleDbCommand();
            command.CommandText = qry;
            command.Connection = con;
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            reader = command.ExecuteReader();
            return reader;
        }

        private void WriteToAll(string namefile, OleDbDataReader reader,StringBuilder getdata)
        {
            string folderLocation = Application.StartupPath + "\\dataExport\\";
            bool folderExists = Directory.Exists(folderLocation);
            bool del = true;
            if (!folderExists)
            {
                Directory.CreateDirectory(Application.StartupPath + "\\dataExport\\");
            }
            StreamWriter myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + namefile + ".txt", append: true);

            try
            {
                while (reader.Read())
                {
                    if (reader["id"].ToString() != "")
                    {
                     //   myFile.WriteLine(String.Format("{0}", reader["info"]));
                     //   myFile.Flush();
                        del = false;
                      //  getdata.Append
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                myFile.WriteLine(String.Format("{0}", reader["info"]));
                myFile.Flush();
                reader.Close();
                myFile.Close();
            }
            if (del) // xoa nhung file trong
            {
                File.Delete(Application.StartupPath + "\\dataExport\\" + namefile + ".txt");
            }
        }

        private void bt_start_Click(object sender, EventArgs e)
        {
            int minYear, maxYear, minMonth, maxMonth, minDay, maxDay;
            string strMonth, strDay;
            progressBar1.Value = 0;
            progressBar1.Maximum = 370;
            progressBar1.Step = 1;
            minYear = 0;
            maxYear = 0;
            minMonth = 1;
            maxMonth = 12;
            minDay = 1;
            maxDay = 31;
            timer.Stop();

            string IniName = Application.StartupPath + @"\connect.ini";
            string[] arrDBname = { "HAMS", "STAR", "MITACOSQL" };
            string[] arrData = { };
            OleDbCommand command = new OleDbCommand();
            OleDbDataReader reader;
            string namefile;
            int sumpropress = 0;
          
            for (int tt = 0; tt < arrDBname.Length; tt++) // so db tong can duyet
            {
                string strtbl = "";
                if (tt == 2)
                {
                    strtbl = GetIniValue(arrDBname[tt], "SERVER", IniName).ToString().Trim();
                }
                else
                {
                    strtbl = GetIniValue(arrDBname[tt], "DATABASE", IniName).ToString().Trim();
                }
                if (strtbl == "")  // truong hop xoa db trong file init
                {
                    continue;
                }
                arrData = strtbl.Split(',');


                for (int ii = 0; ii < arrData.Length; ii++)  // So database con can duyet
                {
                    con.Close();

                    if (tt == 2) // connet sql server
                    {
                        SqlConnection consql = getconsql(arrData[ii]);
                    }
                    else  // connect ms access
                    {
                        con = getcon(arrData[ii]);
                    }
                   
                    if (tt == 1) //Star DB =1
                    {
                        DataSet dsId = GetTableTempId("select * from [data$]", IniName, "ID");  // Table temp get ID
                        DataSet ds = GetTableTemp("select IDX,ORGPOLLINGDATA,Mid(ORGPOLLINGDATA, 15, 4) as [Year],Mid(ORGPOLLINGDATA, 19, 2) as [Month] from T_ORGPOLLINGDATA where Len(ORGPOLLINGDATA)>=33", arrData[ii]); // Table temp get Star va co do dai can thiet
                        string info = "";
                        string dd = "";
                        string hh = "";
                        string mstt = "";
                        string id = "";
                        bool del;
                        string query = "";
                        DataSet dsqry = new DataSet();
                        

                        minYear = Convert.ToInt32(ds.Tables[0].Rows[0][2].ToString()); // lay row dau tien
                        maxYear = Convert.ToInt32(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1][2].ToString()); // lay row cuoi cung
                     // maxMonth = Convert.ToInt32(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1][3].ToString()); // lay thang cuoi cung

                        progressBar1.Value = 0;
                        sumpropress = 370 * (maxYear - minYear + 1);
                        progressBar1.Maximum = sumpropress;

                        for (int i = minYear; i <= maxYear; i++)
                        {
                            for (int j = minMonth; j <= maxMonth; j++)
                            {
                                for (int k = minDay; k <= maxDay; k++)
                                {
                                    strDay = k.ToString();
                                    strMonth = j.ToString();

                                    if (k < 10) // cho du 2 ki tu ngay, thang
                                    {
                                        strDay = "0" + k.ToString();
                                    }
                                    if (j < 10)
                                    {
                                        strMonth = "0" + j.ToString();
                                    }
                            //*******************
                                    // string query = "SELECT  [eventDate] +' '+[eventTime]+',,'+[eventCard]+','+[deviceID]+ ',,' as [info],[eventCard] as id  FROM PubEvent WHERE eventDate = '" + i + "/" + strMonth + "/" + strDay + "'";
                                    // string query = "SELECT  Mid(ORGPOLLINGDATA,15,8)+ ' '+ Mid(ORGPOLLINGDATA,24,6)+',,'+  Mid(ORGPOLLINGDATA,7,8)  + ',00,,' as [info] FROM T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) = '" + i + strMonth + strDay + "'";
                                    // DataSet ds = GetTableTemp("select * from T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) = '20181101'", IniName, "STAR");
                                     query = "SELECT  Mid(ORGPOLLINGDATA,15,8) as [dd], Mid(ORGPOLLINGDATA,24,6) as [hh], Mid(ORGPOLLINGDATA,7,8)  as [mstt],ORGPOLLINGDATA FROM T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) = '" + i + strMonth + strDay + "'";
                                    //  string query = "select * from T_ORGPOLLINGDATA WHERE Mid(ORGPOLLINGDATA,15,8) = '20181101'";
                                     StringBuilder getdata = new StringBuilder();
                                    if (con.State == ConnectionState.Open)
                                    {
                                        con.Close();
                                    }
                                    con = getcon(arrData[ii]);
                                    reader = ExcReader(query, con);
                                    if (reader.HasRows)
                                    {
                                        namefile = i.ToString() + strMonth + strDay;
                                        string folderLocation = Application.StartupPath + "\\dataExport\\";
                                        bool folderExists = Directory.Exists(folderLocation);
                                        del = true;
                                        if (!folderExists)
                                        {
                                            Directory.CreateDirectory(Application.StartupPath + "\\dataExport\\");
                                        }
                                        StreamWriter myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + namefile + ".txt", append: true);
                                        try
                                        {
                                            while (reader.Read())
                                            {
                                                //temp = reader["info"].ToString();
                                                //dd = Microsoft.VisualBasic.Strings.Left(temp,8);
                                                //hh = Microsoft.VisualBasic.Strings.Mid(temp,10,6);
                                                //mstt = Microsoft.VisualBasic.Strings.Mid(temp, 18, 8);

                                                dd = Microsoft.VisualBasic.Strings.Left(reader["dd"].ToString(), 4).ToString() + '/' + Microsoft.VisualBasic.Strings.Mid(reader["dd"].ToString(), 5, 2).ToString() + '/' + Microsoft.VisualBasic.Strings.Right(reader["dd"].ToString(), 2).ToString();
                                                //namefile = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 15, 4).ToString() + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 19, 2).ToString() + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 21, 2).ToString(); ;
                                                hh = Microsoft.VisualBasic.Strings.Left(reader["hh"].ToString(), 2).ToString() + ':' + Microsoft.VisualBasic.Strings.Mid(reader["hh"].ToString(), 3, 2).ToString() + ':' + Microsoft.VisualBasic.Strings.Right(reader["hh"].ToString(), 2).ToString();
                                                mstt = reader["mstt"].ToString();
                                                // len=   Microsoft.VisualBasic.Strings.Len(reader["ORGPOLLINGDATA"].ToString());
                                                // id = Microsoft.VisualBasic.Strings.Trim(Scalar("select MANV  from [data$] where MSTHETU = '" + mstt + "'", getcon(GetIniValue("ID", "DATABASE", IniName).ToString().Trim())));
                                                id = "";
                                                if (Microsoft.VisualBasic.Strings.Len(reader["ORGPOLLINGDATA"].ToString()) >= 33)
                                                {
                                                    for (int rid = 0; rid < dsId.Tables[0].Rows.Count; rid++)  // search ra ID tuong ung ma the tu
                                                    {
                                                        if (mstt == dsId.Tables[0].Rows[rid][2].ToString())
                                                        {
                                                            id = Microsoft.VisualBasic.Strings.Trim(dsId.Tables[0].Rows[rid][0].ToString());
                                                            break;
                                                        }
                                                    }
                                                    //id = mstt;
                                                    //string temp1 = "select MANV  from [data$] where MSTHETU = '" + mstt + "'";
                                                    //OleDbConnection con1 = getcon(GetIniValue("ID", "DATABASE", IniName).ToString().Trim());
                                                    //id = Microsoft.VisualBasic.Strings.Trim(Scalar(temp1, con1));
                                                    if (id != "")
                                                    {
                                                        del = false;
                                                        //   StreamWriter myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + namefile + ".txt", append: true);
                                                        info = dd + " " + hh + ",," + id + ",00,,";
                                                        //  myFile.WriteLine(String.Format("{0}", info));
                                                        // string temp = reader["ORGPOLLINGDATA"].ToString();
                                                        //temp = Microsoft.VisualBasic.Strings.Left(temp, 17) + id + ",00,,";
                                                     //   myFile.WriteLine(String.Format("{0}", info));
                                                     //   myFile.Flush();
                                                        getdata.Append(info + "\r\n");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.ToString());
                                        }
                                        finally
                                        {
                                            myFile.WriteLine(getdata.ToString());
                                            myFile.Flush();
                                            reader.Close();
                                            myFile.Close();
                                        }
                                        //****************************
                                        if (del)
                                        {
                                            File.Delete(Application.StartupPath + "\\dataExport\\" + namefile + ".txt");
                                        }
                                    }
                                    progressBar1.PerformStep();
                                }
                            }
                        }
                        lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[ii] + " !!!");
                    }

                    else if (tt == 0)    // Hams DB =0
                    {
                        minYear = Convert.ToInt16(Microsoft.VisualBasic.Strings.Left(Scalar("SELECT  MIN(eventDate) from PubEvent", con), 4));
                        maxYear = Convert.ToInt16(Microsoft.VisualBasic.Strings.Left(Scalar("SELECT  MAX(eventDate) from pubevent ", con), 4));
                    //  maxMonth = Convert.ToInt16(Microsoft.VisualBasic.Strings.Mid(Scalar("SELECT  MAX(eventDate) from pubevent ", con), 6, 2));
                       
                        progressBar1.Value = 0;
                        sumpropress = 370 * (maxYear - minYear + 1);
                        progressBar1.Maximum = sumpropress;
                        for (int i = minYear; i <= maxYear; i++)
                        {
                            for (int j = minMonth; j <= maxMonth; j++)
                            {
                                for (int k = minDay; k <= maxDay; k++)
                                {
                                    strDay = k.ToString();
                                    strMonth = j.ToString();

                                    if (k < 10)
                                    {
                                        strDay = "0" + k.ToString();
                                    }
                                    if (j < 10)
                                    {
                                        strMonth = "0" + j.ToString();
                                    }

                                    string query = "SELECT  [eventDate] +' '+[eventTime]+',,'+[eventCard]+','+[deviceID]+ ',,' as [info],[eventCard] as id  FROM PubEvent WHERE eventDate = '" + i + "/" + strMonth + "/" + strDay + "'"; ;
                                    con = getcon(arrData[ii]);
                                    StringBuilder getdata = new StringBuilder();
                                    using (con)
                                    {
                                        reader = ExcReader(query, con);
                                        if (reader.HasRows)
                                        {
                                            namefile = i.ToString() + strMonth + strDay;
                                            string folderLocation = Application.StartupPath + "\\dataExport\\";
                                            bool folderExists = Directory.Exists(folderLocation);
                                            bool del = true;
                                            if (!folderExists)
                                            {
                                                Directory.CreateDirectory(Application.StartupPath + "\\dataExport\\");
                                            }
                                            StreamWriter myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + namefile + ".txt", append: true);

                                            try
                                            {
                                                while (reader.Read())
                                                {
                                                    if (reader["id"].ToString() != "")
                                                    {
                                                        //   myFile.WriteLine(String.Format("{0}", reader["info"]));
                                                        //   myFile.Flush();
                                                        del = false;
                                                        getdata.Append(reader["info"] + "\r\n");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show(ex.ToString());
                                            }
                                            finally
                                            {
                                                myFile.WriteLine(getdata.ToString());
                                                myFile.Flush();
                                                reader.Close();
                                                myFile.Close();
                                            }
                                            if (del) // xoa nhung file trong
                                            {
                                                File.Delete(Application.StartupPath + "\\dataExport\\" + namefile + ".txt");
                                            }
                                        }
                                      
                                    }
                                    progressBar1.PerformStep();
                                }
                            }
                        }
                        lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[ii] + " !!!");
                    }

                    else  //if (tt == 7) // Mitaco DB
                    {
                        minYear = Convert.ToInt16(Microsoft.VisualBasic.Strings.Mid(Scalarsql("SELECT  MIN(NgayCham) from CheckInOut", consql), 7, 4));
                        maxYear = Convert.ToInt16(Microsoft.VisualBasic.Strings.Mid(Scalarsql("SELECT  MAX(NgayCham) from CheckInOut ", consql), 7, 4));
                     // maxMonth = Convert.ToInt16(Microsoft.VisualBasic.Strings.Mid(Scalarsql("SELECT  MAX(NgayCham) from CheckInOut ", consql), 4, 2));
                       
                        progressBar1.Value = 0;
                        sumpropress = 370 * (maxYear - minYear + 1);
                        progressBar1.Maximum = sumpropress;
                        for (int i = minYear; i <= maxYear; i++)
                        {
                            for (int j = minMonth; j <= maxMonth; j++)
                            {
                                for (int k = minDay; k <= maxDay; k++)
                                {
                                    strDay = k.ToString();
                                    strMonth = j.ToString();

                                    if (k < 10)
                                    {
                                        strDay = "0" + k.ToString();
                                    }
                                    if (j < 10)
                                    {
                                        strMonth = "0" + j.ToString();
                                    }
                                    string query = "SELECT  NgayCham,GioCham,MaChamCong,MaSoMay FROM CheckInOut WHERE NgayCham = '" + i + "-" + strMonth + "-" + strDay + "'"; ;
                                    string info;
                                    string dd = "";
                                    string hh = "";
                                    StringBuilder getdata = new StringBuilder();
                                    SqlCommand commandsql = new SqlCommand();
                                    SqlDataReader readersql;
                                    consql = getconsql(arrData[ii]);

                                    commandsql.CommandText = query;
                                    commandsql.Connection = consql;
                                    readersql = commandsql.ExecuteReader();
                                    if (readersql.HasRows)
                                    {
                                        namefile = i.ToString() + strMonth + strDay;
                                        string folderLocation = Application.StartupPath + "\\dataExport\\";
                                        bool folderExists = Directory.Exists(folderLocation);
                                        if (!folderExists)
                                        {
                                            Directory.CreateDirectory(Application.StartupPath + "\\dataExport\\");
                                        }
                                        StreamWriter myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + namefile + ".txt", append: true);
                                        try
                                        {
                                            while (readersql.Read())
                                            {
                                                dd = ((DateTime)readersql["NgayCham"]).ToString("yyyy/MM/dd");
                                                hh = ((DateTime)readersql["GioCham"]).ToString("HH:mm:ss");
                                                info = dd + ' ' + hh + ",," + readersql["MaChamCong"].ToString() + ',' + readersql["MaSoMay"].ToString() + ",,";
                                          //      myFile.WriteLine(String.Format("{0}", info));
                                          //      myFile.Flush();
                                                getdata.Append(info+"\r\n");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.ToString());
                                        }
                                        finally
                                        {
                                            myFile.WriteLine(getdata.ToString());
                                            myFile.Flush();
                                            readersql.Close();
                                            myFile.Close();
                                        }
                                    }
                                    progressBar1.PerformStep();
                                } //for 3
                            }  //for 2
                        } // for 1
                        lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[ii] + '\\' + arrDBname[2] + " !!!");
                    }  // end else
                }  // for db con
               // Reset();
               timer.Start();
            } // for db tong
            lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Run.......");
        }

        private void bt_ngay_Click(object sender, EventArgs e)
        {
            timer.Stop();
            string IniName = Application.StartupPath + @"\connect.ini";
            string[] arrDBname = { "STAR", "HAMS", "MITACOSQL" };
            string[] arrData = { };
            string name = dtpicker_date.Value.Date.ToString("yyyMMdd");

            for (int ii = 0; ii < arrDBname.Length; ii++) // so DB tong can duyet
            {
                string strtbl = "";
                if (ii == 2)
                {
                    strtbl = GetIniValue(arrDBname[ii], "SERVER", IniName).ToString().Trim();
                }
                else
                {
                    strtbl = GetIniValue(arrDBname[ii], "DATABASE", IniName).ToString().Trim();
                }

                if (strtbl == "")
                {
                    continue;
                }
                arrData = strtbl.Split(',');
                OleDbCommand command = new OleDbCommand();
                OleDbDataReader reader;
                StreamWriter myFile;
                string query = "";

                for (int i = 0; i < arrData.Length; i++) // so DB con can duyet
                {
                    if (ii == 2) // connet sql server
                    {
                        SqlConnection con = getconsql(arrData[i]);
                    }
                    else  // connect ms access
                    {
                        con = getcon(arrData[i]);
                    }

                    if (ii == 0) // STAR DB
                    {
                        // dataset dsid = gettabletempid("select * from [data$]", ininame, "id");  // table temp get id
                        DataSet ds = GetTableTemp("select idx, orgpollingdata from t_orgpollingdata where len(orgpollingdata)>=33 and mid(orgpollingdata,15,8)='" + name + "'", arrData[i]); // table temp get star
                        string info = "";
                        bool flginfo = false;
                        string dd = "";
                        string hh = "";
                        string id = "";
                        StringBuilder getData = new StringBuilder();

                        for (int r = 0; r < ds.Tables[0].Rows.Count; r++)
                        {
                            dd = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 15, 4).ToString() + '/' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 19, 2).ToString() + '/' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 21, 2).ToString(); ;
                            hh = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 24, 2) + ':' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 26, 2) + ':' + Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 28, 2);
                            id = Microsoft.VisualBasic.Strings.Mid(ds.Tables[0].Rows[r][1].ToString(), 7, 8);
                            //for (int rid = 0; rid < dsId.Tables[0].Rows.Count; rid++)
                            //for (int rid = 0; rid < dsid.tables[0].rows.count; rid++)
                            //{
                            //    if (id == dsid.tables[0].rows[rid][2].tostring())
                            //    {
                            //        id = microsoft.visualbasic.strings.trim(dsid.tables[0].rows[rid][0].tostring());
                            //        break;
                            //    }
                            //}
                            id = Microsoft.VisualBasic.Strings.Trim(Scalar("select MANV  from [data$] where MSTHETU = '" + id + "'", getcon(GetIniValue("ID", "DATABASE", IniName).ToString().Trim())));
                            info = dd + ' ' + hh + ",," + id + ",00,,";
                            getData.Append(info);
                            //if (r < ds.Tables[0].Rows.Count - 1)
                            //{
                                getData.Append("\r\n");
                            //}
                            
                            flginfo = true;
                          }
                            myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                            myFile.WriteLine(getData.ToString());
                            myFile.Flush();
                            myFile.Close();

                        if (flginfo)
                        {
                            lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + " !!!");
                        }
                        else
                        {
                             lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Not data " + arrData[i] + " !!!");
                        }

                    }
                  
                    else if (ii == 1) //HAMS DB
                    {
                       query = "SELECT  [eventDate] +' '+[eventTime]+',,'+[eventCard]+','+[deviceID]+ ',,' as [info]  FROM PubEvent WHERE eventDate = '" + dtpicker_date.Text + "'";
                       StringBuilder getdata = new StringBuilder();
                      using (con)
                        {
                            command.CommandText = query;
                            command.Connection = con;
                            reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {
                                try
                                {
                                    while (reader.Read())
                                    {
                                        //myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                        //myFile.WriteLine(String.Format("{0}", reader["info"]));
                                        //myFile.Flush();
                                        //myFile.Close();
                                      getdata.Append(reader["info"] + "\r\n");
                                      }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    Close();
                                }
                                finally
                                {
                                    myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                    myFile.WriteLine( getdata.ToString());
                                    myFile.Flush();
                                    myFile.Close();
                                    reader.Close();
                                    lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + " !!!");
                                 }
                            }
                            else
                            {
                                lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Not data " + arrData[i] + " !!!");
                            }
                        }
                    }
                    else  //MITACO DB
                    {
                        query = "SELECT  NgayCham,GioCham,MaChamCong,MaSoMay FROM CheckInOut WHERE NgayCham = '" + dtpicker_date.Value.Date.ToString("yyy-MM-dd") + "'";
                        StringBuilder getdata = new StringBuilder();
                        using (consql)
                        {
                            string info;
                            string dd = "";
                            string hh = "";
                            SqlCommand commandsql = new SqlCommand();
                            SqlDataReader readersql;
                            commandsql.CommandText = query;
                            commandsql.Connection = consql;
                            readersql = commandsql.ExecuteReader();
                            if (readersql.HasRows)
                            {
                                try
                                {
                                    while (readersql.Read())
                                    {
                                        dd = ((DateTime)readersql["NgayCham"]).ToString("yyyy/MM/dd");
                                        hh = ((DateTime)readersql["GioCham"]).ToString("HH:mm:ss");
                                        info = dd + ' ' + hh + ",," + readersql["MaChamCong"].ToString() + ',' + readersql["MaSoMay"].ToString() + ",,";
                                        //myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                        //myFile.WriteLine(String.Format("{0}", info));
                                        //myFile.Flush();
                                        //myFile.Close();
                                        getdata.Append(info + "\r\n");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    Close();
                                }
                                finally
                                {
                                    myFile = new StreamWriter(Application.StartupPath + "\\dataExport\\" + name + ".txt", append: true);
                                    myFile.WriteLine(getdata.ToString());
                                    myFile.Flush();
                                    myFile.Close();
                                    readersql.Close();
                                    lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Transfer Successfully " + arrData[i] + '\\' + arrDBname[2] + " !!!");
                                }
                            }
                            else
                            {
                                lbox_info.Items.Add(DateTime.Now.ToShortTimeString() + ", Not data " + arrData[i] + '\\' + arrDBname[2] + " !!!");
                            }
                        }

                    }

                }
                con.Close();
            }
            Reset();
            timer.Start();
        }  
    }
}
