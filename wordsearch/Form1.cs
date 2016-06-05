using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Diagnostics;



namespace wordsearch
{
    
    public partial class Form1 : Form
    {
        string foldername = "C:\\", exname = ".docx",system="C:\\Users" + "\\" + Environment.UserName;//dim location
        string location = System.Environment.CurrentDirectory; 
        string[] wordname = new string[1000];
        string[] wordlocation = new string[1000];
        string[] favlocation = new string[1000];
        string[] favname = new string[1000];
        System.IO.StreamReader filereader;
        System.IO.StreamWriter filewriter;
        int word_index,fav_index=0,flag=1;
        
         
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
         //   MessageBox.Show(system);
            /*Button[] buttons = new Button[3];
            for (int i = 0; i < buttons.Length; i++)
            {
                buttons[i] = new Button();
                buttons[i].Name = "button" + i;
                buttons[i].Text = buttons[i].Name;
                buttons[i].Location = new Point(10, 30 * i);
                buttons[i].Click += new EventHandler(Buttons_Click);
            }
            this.Controls.AddRange(buttons);*/
            FileStream fileStream;
            try
            {
                if (!System.IO.File.Exists(@system + "\\favorite.txt"))
                { fileStream = new FileStream(@system + "\\favorite.txt", FileMode.Create); fileStream.Close(); }
               
                
              // filereader = new System.IO.StreamReader(@location + "\\favorite.txt");
               // filewriter = new System.IO.StreamWriter(@location + "\\favorite.txt");
            }
            catch (Exception x)
            { MessageBox.Show(x.ToString()); }
            label1.Text = wordsearch.Properties.Settings.Default.filelocation;
            foldername = wordsearch.Properties.Settings.Default.filelocation;
            contextMenuStrip1.Items.Add("執行");
            contextMenuStrip1.Items.Add("加入我的最愛");
            contextMenuStrip1.Items.Add("從我的最愛刪除");
            load_fav();
            printlist(1);
        }
        void Buttons_Click(object sender, EventArgs e)
        {
            this.Text = (sender as Button).Text;
        }

        void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBox1.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                if (flag == 0)
                { Process.Start(wordlocation[index]); }
                else if (flag == 1)
                { Process.Start(favlocation[index]); }
                //do your stuff here
            }
        }
        void printlist(int tag)
        {
            DataTable dt = new DataTable();
                
                dt.Columns.Add("word");
                if (tag == 0)
                {
                    for (int i = 0; i < word_index; i++)
                    {
                        dt.Rows.Add("(" + (i + 1) + ") " + wordname[i]);
                    }
                }
                else
                {
                    for (int i = 0; i < fav_index; i++)
                    {
                        dt.Rows.Add("(" + (i + 1) + ") " + favlocation[i]);
                    }
                }
            dt.AcceptChanges();
            this.listBox1.DisplayMember = "word";
            this.listBox1.DataSource = dt;
            this.listBox1.Height = 200; 
            this.listBox1.MouseDoubleClick += new MouseEventHandler(listBox1_MouseDoubleClick);
        }
        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            label1.Text = wordlocation[this.listBox1.SelectedIndex];         
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                flag = 0;
                foldername = wordsearch.Properties.Settings.Default.filelocation;
                searchfile();
                printlist(0);
            }
            catch (Exception x)
            {
               MessageBox.Show(x.ToString());
            }
        }
        private void searchfile()
        {
            List<string> smaliList = new List<string>();
         try
            {
                // Path.GetFileNameWithoutExtension(filelocation);
                word_index = 0;
                ListFiles(new DirectoryInfo(foldername));
            }
            catch (Exception x)
            {
                MessageBox.Show("Error Searchfile\r\n" + x.Message);
                return;
            }

        }
        public void ListFiles(FileSystemInfo info)
        {
            
            if (!info.Exists) return;
            DirectoryInfo dir =info as DirectoryInfo;

         /* DirectorySecurity ds = dir.GetAccessControl();
               FileSystemAccessRule ar1 = new FileSystemAccessRule(Environment.UserDomainName + "\\" + Environment.UserName
                   , FileSystemRights.Read, AccessControlType.Allow);
               FileSystemAccessRule ar2 = new FileSystemAccessRule(Environment.UserDomainName + "\\" + Environment.UserName, 
                   FileSystemRights.Read, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.InheritOnly, AccessControlType.Allow);
             //  ds.AddAccessRule(ar1);
           //    ds.AddAccessRule(ar2);
               dir.SetAccessControl(ds);*/
            /*DirectorySecurity ds = Directory.GetAccessControl(foldername, AccessControlSections.All);
            ds.AddAccessRule(new FileSystemAccessRule(Environment.UserDomainName + "\\" + Environment.UserName,
                                   FileSystemRights.FullControl,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                   PropagationFlags.None,
                                   AccessControlType.Allow));
            Directory.SetAccessControl(foldername, ds);*/
            if (dir == null) return;
            FileSystemInfo[] files = dir.GetFileSystemInfos();           
            //---------------------------------------------
            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = files[i] as FileInfo;
                //是文件 
                if (file != null && file.Extension == exname)
                {
                   
                    // MessageBox.Show(file.FullName);
                    wordlocation[word_index] = file.FullName;
                    wordname[word_index] =Path.GetFileNameWithoutExtension(wordlocation[word_index]);                    
                     word_index++;
                    //對於子目錄，進行遞歸調用 
                }
                else
                    ListFiles(files[i]);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Tag = "選擇檔案";
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    wordsearch.Properties.Settings.Default.filelocation = folderBrowserDialog1.SelectedPath;
                    wordsearch.Properties.Settings.Default.Save();
                    foldername = wordsearch.Properties.Settings.Default.filelocation;
                    label1.Text = foldername;
                }
            else if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                {                  
                }
          
        }
        public void ChangeAuthorities(string destinationPath, FileAttributes fileAttributes)
        {
            // 修改檔案權限為可讀寫
            if (Directory.Exists(destinationPath))
            {
                DirectoryInfo dir = new DirectoryInfo(destinationPath);
                foreach (var file in dir.GetFiles())
                {
                    file.Attributes = fileAttributes;
                }
                foreach (var item in dir.GetDirectories())
                {
                    ChangeAuthorities(item.FullName, fileAttributes);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           // MessageBox.Show(comboBox1.SelectedItem.ToString());
            switch (comboBox1.SelectedItem.ToString())
                {
                    case "word":
                        exname = ".docx";
                        break;
                    case "ppt":
                        exname = ".pptx";
                        break;
                    case "excel":
                        exname=".xlsx";
                        break;
                    case "pdf":
                         exname=".pff";
                        break;
                    case "自訂":
                         exname="."+textBox1.Text;
                        // MessageBox.Show(exname);
                        break;
                    default:
                        exname=".docx";
                        break;
                }
        }

        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            //int height = 0;
          // MessageBox.Show(listBox1.Items.Count.ToString());
            if (e.Button == MouseButtons.Right)
            { // MessageBox.Show("d");
                listBox1.SelectedIndex = listBox1.IndexFromPoint(e.X, e.Y);
                //int currentindex = e.Y / 20;
                //listBox1.SetSelected(currentindex, true);
                contextMenuStrip1.Show(MousePosition);            
            }
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
           // filewriter = File.AppendText(@location + "\\favorite.txt");
           // filewriter = new System.IO.StreamWriter(@location + "\\favorite.txt");
            if (e.ClickedItem.ToString() == "加入我的最愛")
            {
                bool same=false;
               // filewriter.WriteLine(wordlocation[listBox1.SelectedIndex]);
                for (int i = 0; i < fav_index; i++)
                {
                    if (wordlocation[listBox1.SelectedIndex] == favlocation[i])
                    { 
                        same = true;
                    break;
                    }
                }
                if (!same)
                { favlocation[fav_index] = wordlocation[listBox1.SelectedIndex];
                fav_index++;
               


                }
               
            }
            if (e.ClickedItem.ToString() == "執行")
            {
                Process.Start(favlocation[listBox1.SelectedIndex]);
            }
            if (e.ClickedItem.ToString() == "從我的最愛刪除" && flag==1)
            {
                for (int i = listBox1.SelectedIndex; i < fav_index-1; i++)
                {
                    string temp = favlocation[i + 1];
                    favlocation[i] = temp;
                }
                fav_index--;
                printlist(1);
            }
            //filewriter.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {          
            flag = 1;            
            printlist(1);
        }
        void load_fav()
        {
            filereader = new System.IO.StreamReader(@system + "\\favorite.txt");
            fav_index = 0;
            while (!filereader.EndOfStream )
            {

                   
                    favlocation[fav_index] = filereader.ReadLine();
                   // MessageBox.Show(favlocation[fav_index]);
                    fav_index++;
                
            }
            filereader.Close();
        }
        void show_fav()
        {
             
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            filewriter = new System.IO.StreamWriter(@system + "\\favorite.txt");
            for (int i = 0; i < fav_index; i++)
            {
                filewriter.WriteLine(favlocation[i] );
            }
            filewriter.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox1.SelectedText = "自訂";
            comboBox1.SelectedItem = "自訂";
            exname = "."+textBox1.Text; 
        }
     
    }
}
