using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.Threading;
using System.Text;

namespace FTE_EngineCEA
{
    class main
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());
        }
    }

    class Window : Form
    {
        //文字列
        const int itemNum = 11;
        int[] fontSize = new int[] { 13, 16 };
        Label[] item = new Label[itemNum];

        //ボタン
        const int buttonNum = 5;
        Button[] button = new Button[buttonNum];

        //ツールチップ
        const int toolTipNum = 5;
        ToolTip[] toolTip = new ToolTip[toolTipNum];

        //ラジオボタン
        RadioButton[,] radioButton = new RadioButton[4, 2];

        //グループボックス
        const int groupNum = 4;
        GroupBox[] groupBox = new GroupBox[groupNum];

        //テキストボックス
        const int textBoxNum = 20;
        TextBox[] textBox = new TextBox[textBoxNum];

        //デフォルトの値
        struct defaultValue
        {
            public string burningTemp;
            public string OxName, OxFormula, OxTemp, OxAmount, OxEneH;
            public string FuName, FuFormula, FuTemp, FuAmount, FuEneH;
        } defaultValue defValue;

        //初期O/F,燃焼室圧力,C*,γ
        struct initialValue
        {
            public double OF, comPres, Cstar, gamma;
            public int OFnum;
        }initialValue iniValue;
        //テンプレシートの初期タンク圧のやつ
        int maxComPresNum = -1;
        double maxComPres = -1;
        double templateInitialTankPres;

        //dllインポート
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        /*[DllImport("KERNEL32.DLL")]
        public static extern int
          GetPrivateProfileString(string lpAppName,
          string lpKeyName, string lpDefault,
          StringBuilder lpReturnedString, int nSize,
          string lpFileName);*/

        public Window()//GUI基本構造
        {
            //ウィンドウ
            this.Width = 960;
            this.Height = 430;
            this.Text = "FTE-EngineCEA";
            try
            {
                this.Icon = new System.Drawing.Icon("Other\\FTE.ico");
            }
            catch{ }

            //settingファイル
            string settingfilePath = "Other\\DefaultSettings.txt";
            readSettings("ChamberTemperature", ref defValue.burningTemp, settingfilePath);
            readSettings("OxidName", ref defValue.OxName, settingfilePath);
            readSettings("OxidAmount", ref defValue.OxAmount, settingfilePath);
            readSettings("OxidTemp", ref defValue.OxTemp, settingfilePath);
            readSettings("OxidFormula", ref defValue.OxFormula, settingfilePath);
            readSettings("OxidEnergyH", ref defValue.OxEneH, settingfilePath);
            readSettings("FuelName", ref defValue.FuName, settingfilePath);
            readSettings("FuelAmount", ref defValue.FuAmount, settingfilePath);
            readSettings("FuelTemp", ref defValue.FuTemp, settingfilePath);
            readSettings("FuelFormula", ref defValue.FuFormula, settingfilePath);
            readSettings("FuelEnergyH", ref defValue.FuEneH, settingfilePath);


            //初期化
            for (int i = 0; i < itemNum; i++)
            {
                this.item[i] = new Label();
                this.item[i].AutoSize = true;
            }
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    this.radioButton[i, j] = new RadioButton();
                    this.radioButton[i, j].AutoSize = true;
                }
            }
            for (int i = 0; i < buttonNum; i++)
            {
                this.button[i] = new Button();
            }
            for (int i = 0; i < toolTipNum; i++)
            {
                this.toolTip[i] = new ToolTip();
            }
            for (int i = 0; i < textBoxNum; i++)
            {
                this.textBox[i] = new TextBox();
                this.textBox[i].AutoSize = true;
            }
            for (int i = 0; i < groupNum; i++)
            {
                this.groupBox[i] = new GroupBox();
            }
            this.radioButton[0, 0].Checked = true;
            this.radioButton[1, 0].Checked = true;
            this.radioButton[2, 0].Checked = true;
            this.radioButton[3, 1].Checked = true;
            radioButtonManager(1, 1);
            radioButtonManager(2, 1);

            //ラジオボタングループ化
            RadioButton[] rb = new RadioButton[2];
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    rb[j] = this.radioButton[i, j];
                }
                this.groupBox[i].Controls.AddRange(rb);
            }

            //ラジオボタン
            this.groupBox[0].Size = new System.Drawing.Size(125, 40);
            this.groupBox[0].Location = new System.Drawing.Point(90, 40);
            this.groupBox[1].Size = new System.Drawing.Size(125, 40);
            this.groupBox[1].Location = new System.Drawing.Point(250, 80);
            this.groupBox[2].Size = new System.Drawing.Size(125, 40);
            this.groupBox[2].Location = new System.Drawing.Point(80, 200);
            this.groupBox[3].Size = new System.Drawing.Size(125, 40);
            this.groupBox[3].Location = new System.Drawing.Point(80, 240);

            //項目
            this.item[0].Location = new System.Drawing.Point(10, 5);
            this.item[0].Font = new System.Drawing.Font("メイリオ", fontSize[1], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[0].Text = "nasa-cea設定:Rocket-rkt・Frozen・Infinite Area・Throat・Estimate";

            this.item[1].Location = new System.Drawing.Point(30, 50);
            this.item[1].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[1].Text = "O/F:";

            this.item[4].Location = new System.Drawing.Point(30, 90);
            this.item[4].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[4].Text = "燃焼圧[bar(=0.1MPa)]:";

            for (int i = 0; i < 2; i++)
            {
                this.radioButton[i, 0].Location = new System.Drawing.Point(10, 15);
                this.radioButton[i, 0].Text = "範囲";
                this.radioButton[i, 1].Location = new System.Drawing.Point(70, 15);
                this.radioButton[i, 1].Text = "指定";
            }

            this.item[7].Location = new System.Drawing.Point(30, 130);
            this.item[7].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[7].Text = "燃焼室温度[K]:";
            this.textBox[8].Location = new System.Drawing.Point(180, 135);
            this.textBox[8].Text = defValue.burningTemp;

            this.item[8].Location = new System.Drawing.Point(30, 170);
            this.item[8].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[8].Text = "反応物質: 種類　　 　　　名前　　　  　　質量比[%]　 　 　温度[K]　　　  　化学式               EnergyH";

            this.textBox[9].Location = new System.Drawing.Point(230, 215);
            this.textBox[9].Text = defValue.OxName;
            this.textBox[10].Location = new System.Drawing.Point(380, 215);
            this.textBox[10].Text = defValue.OxAmount;
            this.textBox[11].Location = new System.Drawing.Point(530, 215);
            this.textBox[11].Text = defValue.OxTemp;
            this.textBox[12].Location = new System.Drawing.Point(680, 215);
            this.textBox[12].Text = defValue.OxFormula;
            this.textBox[13].Location = new System.Drawing.Point(830, 215);
            this.textBox[13].Text = defValue.OxEneH;

            this.textBox[14].Location = new System.Drawing.Point(230, 255);
            this.textBox[14].Text = defValue.FuName;
            this.textBox[15].Location = new System.Drawing.Point(380, 255);
            this.textBox[15].Text = defValue.FuAmount;
            this.textBox[16].Location = new System.Drawing.Point(530, 255);
            this.textBox[16].Text = defValue.FuTemp;
            this.textBox[17].Location = new System.Drawing.Point(680, 255);
            this.textBox[17].Text = defValue.FuFormula;
            this.textBox[18].Location = new System.Drawing.Point(830, 255);
            this.textBox[18].Text = defValue.FuEneH;

            for (int i = 2; i < 4; i++)
            {
                this.radioButton[i, 0].Location = new System.Drawing.Point(10, 15);
                this.radioButton[i, 0].Text = "酸化剤";
                this.radioButton[i, 1].Location = new System.Drawing.Point(70, 15);
                this.radioButton[i, 1].Text = "燃料";
            }

            this.item[10].Location = new System.Drawing.Point(30, 300);
            this.item[10].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
            this.item[10].Text = "ファイル名:";
            this.textBox[19].Location = new System.Drawing.Point(160, 305);
            this.textBox[19].Text = "";

            //ボタン
            this.button[0].Location = new System.Drawing.Point(100, 350);
            this.button[0].Text = "inp出力";
            this.button[1].Location = new System.Drawing.Point(250, 350);
            this.button[1].Text = "cea起動";
            this.button[2].Location = new System.Drawing.Point(400, 350);
            this.button[2].Text = "Excel出力";
            this.button[3].Location = new System.Drawing.Point(550, 350);
            this.button[3].Text = "個別計算";
            this.button[4].Location = new System.Drawing.Point(700, 350);
            this.button[4].Text = "Excelを開く";

            //ツールチップ
            toolTipInitialize(toolTip[0], button[0], "ファイル名.inpを出力→cea起動へ");
            toolTipInitialize(toolTip[1], button[1], "ファイル名.outを出力→excel出力へ");
            toolTipInitialize(toolTip[2], button[2], "ceaで計算した結果をexcelへファイル名.xlsxとして出力");
            toolTipInitialize(toolTip[3], button[3], "O/Fと燃焼室圧力が1つずつ指定された場合のただ1つの結果を表示したい時に使用");
            toolTipInitialize(toolTip[4], button[4], "ファイル名.xlsxを開く");

            //イベント
            this.button[0].Click += new EventHandler(saveINP);
            this.button[1].Click += new EventHandler(runCEA);
            this.button[2].Click += new EventHandler(saveAsXLSX);
            this.button[3].Click += new EventHandler(calcIndividually);
            this.button[4].Click += new EventHandler(openFileByExcel);
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    this.radioButton[i, j].Click += new EventHandler(radioButtonState);
                    this.radioButton[i, j].Click += new EventHandler(radioButtonState);
                }
            }

            //描画
            for (int i = 0; i < itemNum; i++)
                this.Controls.Add(this.item[i]);
            for (int i = 0; i < textBoxNum; i++)
                this.Controls.Add(this.textBox[i]);
            for (int i = 0; i < groupNum; i++)
                this.Controls.Add(this.groupBox[i]);
            for (int i = 0; i < buttonNum; i++)
                this.Controls.Add(this.button[i]);

        }

        void saveINP(object sender, EventArgs e)//.inpの保存 run by button
        {
            //ファイル
            string fileName = this.textBox[19].Text + ".inp";
            string filePath = System.IO.Path.Combine("NASA-CEA\\" + fileName);
            string text;
            //例外
            int Switch = 1;
            double Temp;

            if (this.radioButton[0, 0].Checked == true)
            {//OF範囲指定
                if (this.textBox[0].Text == "" || this.textBox[1].Text == "" || this.textBox[3].Text == ""
                || double.TryParse(this.textBox[0].Text, out Temp) == false || double.TryParse(this.textBox[1].Text, out Temp) == false || double.TryParse(this.textBox[3].Text, out Temp) == false)
                {
                    Switch = -1;
                    MessageBox.Show("O/Fの値が誤っています", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (double.TryParse(this.textBox[0].Text, out Temp) == true && double.TryParse(this.textBox[1].Text, out Temp) == true)
                {
                    if (double.Parse(this.textBox[0].Text) > double.Parse(this.textBox[1].Text)
                        || double.TryParse(this.textBox[3].Text, out Temp) == true && (double.Parse(this.textBox[1].Text) - double.Parse(this.textBox[0].Text) < double.Parse(this.textBox[3].Text)))
                    {
                        Switch = -1;
                        MessageBox.Show("O/Fの範囲に誤りがあります", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (this.textBox[3].Text == "0") 
                {
                    Switch = -1;
                    MessageBox.Show("間隔を0にすることは出来ません。「指定」オプションを使用してください。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else//OF個別指定
            {
                if (this.textBox[2].Text == "")
                {
                    Switch = -1;
                    MessageBox.Show("O/Fの値が誤っています", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (this.radioButton[1, 0].Checked == true)
            {//燃焼圧範囲指定
                if (this.textBox[4].Text == "" || this.textBox[5].Text == "" || this.textBox[7].Text == ""
                   || double.TryParse(this.textBox[4].Text, out Temp) == false || double.TryParse(this.textBox[5].Text, out Temp) == false || double.TryParse(this.textBox[7].Text, out Temp) == false)
                {
                    Switch = -1;
                    MessageBox.Show("燃焼圧の値が誤っています", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (double.TryParse(this.textBox[4].Text, out Temp) == true && double.TryParse(this.textBox[5].Text, out Temp) == true)
                {
                    if (double.Parse(this.textBox[4].Text) > double.Parse(this.textBox[5].Text)
                        || double.TryParse(this.textBox[7].Text, out Temp) == true && (double.Parse(this.textBox[5].Text) - double.Parse(this.textBox[4].Text) < double.Parse(this.textBox[7].Text)))
                    {
                        Switch = -1;
                        MessageBox.Show("燃焼圧の範囲に誤りがあります", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (this.textBox[3].Text == "0")
                {
                    Switch = -1;
                    MessageBox.Show("間隔を0にすることは出来ません。「指定」オプションを使用してください。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else//燃焼圧個別指定
            {
                if (this.textBox[6].Text == "")
                {
                    Switch = -1;
                    MessageBox.Show("燃焼圧の値が誤っています", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (double.TryParse(this.textBox[8].Text, out Temp) == false)
            {
                Switch = -1;
                MessageBox.Show("燃焼室温度の値に誤りがあります", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            switch (Switch)
            {
                case 1:
                    text = "problem    o/f=";
                    if (this.radioButton[0, 0].Checked == true)//OF範囲指定
                    {
                        double last = double.Parse(this.textBox[1].Text);
                        double first = double.Parse(this.textBox[0].Text);
                        double interval = double.Parse(this.textBox[3].Text);
                        int OFNum = (int)(Math.Round(((last - first) / interval) + 1));
                        double[] value = new double[OFNum];
                        for (int i = 0; i < OFNum; i++)//はじめと終わりと間隔からほしい数値を取得してカンマで区切る
                        {
                            value[i] = double.Parse(this.textBox[0].Text) + i * double.Parse(this.textBox[3].Text);
                            text += value[i].ToString() + ",";
                        }
                    }
                    else//OF個別指定
                    {
                        text += this.textBox[2].Text.ToString();
                        if (this.textBox[2].Text[this.textBox[2].Text.Length - 1] != ',') 
                            text += ',';
                    }
                    text += Environment.NewLine + "      rocket  frozen  nfz=2  tcest,k=" + this.textBox[8].Text + Environment.NewLine + "  p,bar=";

                    if (this.radioButton[1, 0].Checked == true)//燃焼圧範囲指定
                    {
                        double last = double.Parse(this.textBox[5].Text);
                        double first = double.Parse(this.textBox[4].Text);
                        double interval = double.Parse(this.textBox[7].Text);
                        int BarNum = (int)(Math.Round(((last - first) / interval) + 1));
                        double[] value = new double[BarNum];
                        for (int i = 0; i < BarNum; i++)//はじめと終わりと間隔からほしい数値を取得してカンマで区切る
                        {
                            value[i] = double.Parse(this.textBox[4].Text) + i * double.Parse(this.textBox[7].Text);
                            text += value[i].ToString() + ",";
                        }
                    }
                    else//燃焼圧個別指定
                    {
                        text += this.textBox[6].Text.ToString();
                        if (this.textBox[6].Text[this.textBox[6].Text.Length - 1] != ',')
                            text += ',';
                    }
                    text += Environment.NewLine + "react" + Environment.NewLine;
                    if (this.radioButton[2, 0].Checked == true)
                        text += " oxid=";
                    else
                        text += " fuel=";
                    text += this.textBox[9].Text + " wt=" + this.textBox[10].Text + "  t,k=" + this.textBox[11].Text + Environment.NewLine;
                    if (this.radioButton[3, 0].Checked == true)
                        text += " oxid=";
                    else
                        text += " fuel=";
                    text += this.textBox[14].Text + " wt=" + this.textBox[15].Text + "  t,k=" + this.textBox[16].Text;
                    text += Environment.NewLine + "    h,kj/mol=" + this.textBox[18].Text + "  ";

                    //化学式を分割
                    char GetCF = this.textBox[17].Text[0];
                    int swit = 0;//0:文字 1:数字
                    for (int i = 0; i < this.textBox[17].Text.Length; i++)
                    {
                        GetCF = this.textBox[17].Text[i];
                        if (char.IsNumber(GetCF) == true)//整数だったら
                        {
                            if (swit == 0)//その前が整数以外だったら
                            {
                                text += " ";//空白を追加
                                swit = 1;
                            }
                            text += GetCF;
                        }
                        else//整数以外（アルファベット）だったら
                        {
                            if (swit == 1)//その前が整数だったら
                            {
                                text += " ";//空白を追加
                                swit = 0;
                            }
                            text += GetCF;
                        }
                    }
                    text += Environment.NewLine + "output  short" + Environment.NewLine + "end";

                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(filePath))
                    {
                        file.Write(text);
                    }
                    MessageBox.Show(this.textBox[19].Text + ".inpを保存しました", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    System.Threading.Thread.Sleep(200);
                    break;

                case -1:
                    break;
            }
        }

        void runCEA(object sender, EventArgs e)//FCEA2mの起動 run by button
        {
            try
            {
                System.Diagnostics.Process.Start("NASA-CEA\\FCEA2m.exe");
                System.Threading.Thread.Sleep(50);
                foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
                {
                    //"FCEA2m"がメインウィンドウのタイトルに含まれているか調べる
                    if (0 <= p.MainWindowTitle.IndexOf("FCEA2m"))
                    {
                        //ウィンドウをアクティブにし、「NASA-CEA\\inpファイル名」を自動入力後エンター
                        SetForegroundWindow(p.MainWindowHandle);
                        System.Threading.Thread.Sleep(50);
                        SendKeys.Send("NASA-CEA\\" + this.textBox[19].Text);
                        if (this.textBox[19].Text != "")
                            SendKeys.Send("{ENTER}");
                        System.Threading.Thread.Sleep(50);
                        break;
                    }
                }
            }
            catch
            {
                MessageBox.Show("NASA-CEAを起動できません。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        int getReadingIndex(string valueName, string source, int dataStartingNum = 0)//読み取る文字の開始位置を取得(.out内) in saveAsXLSX(),inputDatasToCells()
        {
            char GetData;
            int count = 0;
            int GetCounter = 0;
            string judgement = "";
            for (int i = dataStartingNum; i < source.Length; i++)
            {
                GetData = source[i];//位置文字ずつ見てく
                for (int j = 0; j < valueName.Length + 1; j++) 
                {
                    if (count == 0)
                        judgement = "";
                    if (count == valueName.Length - j)
                    {
                        if (GetData == valueName[count])//取得した文字がほしい文字のcount番目と一致
                        {
                            judgement += GetData;
                            count += 1;
                            if (count == valueName.Length)
                            {
                                GetCounter += 1;
                                count = 0;
                                if (judgement == valueName)//欲しいデータの場所まで来てかつ文字列があってる
                                {
                                    judgement = "";
                                    return i;//現在地を返す
                                }
                            }
                        }
                        else count = 0;
                    }

                }
            }
            return 0;
        }

        void saveAsXLSX(object sender, EventArgs e)//.outの内容を.xlsxで保存 run by button
        {
            //例外処理
            int Switch = 0;
            //ファイル
            string File_Name = this.textBox[19].Text + ".out";
            string file_path = System.IO.Path.Combine("NASA-CEA\\" + File_Name);
            System.IO.StreamReader file;
            string OUTfileAllData = "";//.outの全データ格納
            //処理系
            int equalNum = 0;
            bool newLine = false;
            //.out読み込み
            try{
                file = new System.IO.StreamReader(file_path);
                OUTfileAllData = file.ReadToEnd();
                file.Close();
                Switch = 1;
            }
            catch
            {
                Switch = -1;
            }
            switch (Switch) {
                case 1:
                    //OF読み込み
                    equalNum = getReadingIndex("o/f=", OUTfileAllData);//=の位置を取得
                    char getChar;
                    int commaNum = 0;
                    int endValueNum = 0;
                    for (int i = equalNum + 1; i < OUTfileAllData.Length; i++)
                    {
                        getChar = OUTfileAllData[i];
                        if (getChar == ',') commaNum += 1;//コンマが何個あるか→データ数の取得
                        if (getChar == '\n')
                        {
                            i += 2;
                            getChar = OUTfileAllData[i];
                            newLine = true;
                        }
                        if (getChar != ',' && newLine == true) //改行かつOFデータは終わり
                        {
                            endValueNum = i;
                            break;
                        }
                        if (newLine == true)
                        {
                            commaNum++;
                            newLine = false;
                        }
                    }

                    string[] OFData = new string[commaNum];
                    int ValueCounter = 0;

                    for (int i = equalNum + 1; i < endValueNum; i++)
                    {
                        getChar = OUTfileAllData[i];
                        if (newLine == true && getChar == '\n')
                        {
                            OFData[ValueCounter] = OFData[ValueCounter].Trim();
                            i += 2;
                            getChar = OUTfileAllData[i];
                        }
                        if (getChar != ',')
                            OFData[ValueCounter] += getChar;//数値を取得
                        else ValueCounter += 1;
                        if (ValueCounter == commaNum) break;//全部入れたら終わり
                    }

                    //bar読み込み
                    equalNum = getReadingIndex("p,bar=", OUTfileAllData);
                    commaNum = 0;
                    endValueNum = 0;
                    for (int i = equalNum + 1; i < OUTfileAllData.Length; i++)
                    {
                        getChar = OUTfileAllData[i];
                        if (getChar == ',') commaNum += 1;
                        if (getChar == '\n')
                        {
                            endValueNum = i;
                            break;
                        }
                    }

                    string[] BarData = new string[commaNum];
                    for (int i = 0; i < commaNum; i++) BarData[i] = "";
                    ValueCounter = 0;

                    for (int i = equalNum + 1; i < endValueNum; i++)
                    {
                        getChar = OUTfileAllData[i];
                        if (getChar != ',')
                            BarData[ValueCounter] += getChar;
                        else ValueCounter += 1;
                        if (ValueCounter == commaNum) break;
                    }

                    //.xlsx保存
                    //初期化
                    string CurrentDir = System.IO.Directory.GetCurrentDirectory();
                    var workbook = new XLWorkbook();

                    //いっぱいあるデータを取得してセルに入れる
                    makeCalcTemplate(workbook, 1);
                    inputDatasToCells(OUTfileAllData, "CSTAR, M/SEC", "CSTAR", workbook, OFData, BarData, '\n');
                    inputDatasToCells(OUTfileAllData, " GAMMAs            ", "GAMMA", workbook, OFData, BarData, ' ');
                    makeCalcTemplate(workbook, 2);

                    //保存
                    try
                    {
                        workbook.SaveAs(CurrentDir + "\\Simulator\\" + this.textBox[19].Text + ".xlsx");
                        MessageBox.Show(this.textBox[19].Text + ".xlsxを保存しました", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch
                    {
                        MessageBox.Show("保存できません\n開いているファイルを閉じて再度保存してください", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case -1:
                    MessageBox.Show(this.textBox[19].Text + ".outが見つかりません", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        void inputToCells(int XNum, int YNum, string Input, IXLWorksheet ws)//セルに数値を入れる in inputDatasToCells(),makeCalcTemplate()
        {
            ws.Cell(YNum, XNum).Value = Input;
        }

        void makeCalcTemplate(XLWorkbook workbook, int scene)//テンプレ計算シート作成 in saveAsXLSX()
        {
            try
            {
                if (scene == 1)
                {
                    var templateWorkbook = new XLWorkbook("Other\\SimulationTemplate.xlsx");
                    templateWorkbook.Worksheet(1).CopyTo(workbook, "Calculation", 1);
                    var Sheet = workbook.Worksheets.Worksheet("Calculation");
                    //初期タンク圧取得
                    templateInitialTankPres = Sheet.Cell(12, 3).GetDouble();
                }
                if (scene == 2)
                {
                    var Sheet = workbook.Worksheets.Worksheet("Calculation");
                    //自動入力
                    inputToCells(3, 2, iniValue.OF.ToString(), Sheet);
                    inputToCells(3, 5, iniValue.comPres.ToString(), Sheet);
                    inputToCells(3, 3, iniValue.Cstar.ToString(), Sheet);
                    inputToCells(3, 4, iniValue.gamma.ToString(), Sheet);
                }
            }catch
            {
                MessageBox.Show("計算シートを生成できません", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        void inputDatasToCells(string source, string valueName, string sheetName, XLWorkbook workbook, string[] OFData, string[] BarData, char moveToNextWord)//Of,Barまとめてセルに入れる, in saveAsXLSX
        {
            var sheet = workbook.Worksheets.Add(sheetName);
            //データ名をセルに入れる
            inputToCells(2, 1, "O/F", sheet);
            inputToCells(1, 2, "燃焼室圧力[bar]", sheet);

            //値
            //O/F
            for (int i = 0; i < OFData.Length; i++)
            {
                inputToCells(2, 3 + i, OFData[i], sheet);
            }
            //燃焼圧[bar]
            for (int i = 0; i < BarData.Length; i++)
            {
                inputToCells(3 + i, 2, BarData[i], sheet);
            }

            char getc;
            int valueFirst = 0, vCounter = 0; ;
            int valueCount = 1;//何個目のデータがほしいか
            string[] valueData = new string[OFData.Length * BarData.Length];//これに数値を入れてく
            for (int i = 0; i < OFData.Length; i++)
            {
                for (int j = 0; j < BarData.Length; j++)
                {
                    valueFirst = getReadingIndex(valueName, source, valueFirst) + 1;//ほしい数値の位置を取得
                    getc = source[valueFirst];//1文字ずつ見てく
                    while (getc != moveToNextWord)//次のデータに移りたい時に見る文字と比較(スペースとか改行とか)
                    {
                        getc = source[valueFirst + vCounter];
                        valueData[valueCount - 1] += getc;
                        vCounter += 1;
                    }
                    vCounter = 0;
                    valueCount += 1;
                }

            }

            //自動入力の数値計算とか
            if (sheetName == "CSTAR")
            {
                maxComPresNum = -1;
                maxComPres = -1;
                iniValue.Cstar = 0;
                iniValue.OFnum = 0;
            }
            for (int i = 0; i < BarData.Length; i++)//燃焼室圧はタンク圧/2以下
            {
                if (maxComPres < double.Parse(BarData[i]) && double.Parse(BarData[i]) <= templateInitialTankPres * 10 / 2) 
                {
                    maxComPresNum = i;
                    maxComPres = double.Parse(BarData[i]);
                }
            }
            for (int i = 0; i < OFData.Length; i++)
            {
                for (int j = 0; j < BarData.Length; j++)
                {
                    inputToCells(3 + j, 3 + i, valueData[i * BarData.Length + j], sheet);

                    if (sheetName == "CSTAR")//初期値決定　初期燃焼室圧は初期タンク圧/2以下の最大でcstar最大　OF少し大きめ
                    {
                        if (j == maxComPresNum)
                        {
                            if (iniValue.Cstar < double.Parse(valueData[i * BarData.Length + j]))
                            {
                                iniValue.Cstar = double.Parse(valueData[i * BarData.Length + j]);
                                iniValue.OFnum = i;
                            }
                        }
                    }
                    if (sheetName == "GAMMA")
                    {
                        if (i == iniValue.OFnum && j == maxComPresNum)
                        {
                            iniValue.gamma = double.Parse(valueData[i * BarData.Length + j]);
                            iniValue.OF = double.Parse(OFData[iniValue.OFnum]);
                            iniValue.comPres = double.Parse(BarData[maxComPresNum]) / 10;//[Bar] to [MPa]
                        }
                    }
                }
            }

            if (sheetName == "CSTAR")
            {
                if (iniValue.OFnum < OFData.Length - 1) 
                {
                    iniValue.OFnum += 1;
                    iniValue.Cstar = double.Parse(valueData[BarData.Length * iniValue.OFnum + maxComPresNum]); 
                }
            }


            //計算条件を入力
            int extraInfoNum = 0;
            int readingIndex;
            string CFtemp;//化学式計算用
            string inputDataValue = "";
            inputToCells(BarData.Length + 4, 1, "燃焼室温度[K]", sheet);
            inputToCells(BarData.Length + 4, 3, "種類", sheet);
            inputToCells(BarData.Length + 5, 3, "名前", sheet);
            inputToCells(BarData.Length + 6, 3, "質量比[%]", sheet);
            inputToCells(BarData.Length + 7, 3, "温度[K]", sheet);
            inputToCells(BarData.Length + 8, 3, "EnergyH[kJ/mol]", sheet);
            inputToCells(BarData.Length + 9, 3, "化学式", sheet);

            //燃焼室温度
            readingIndex = getReadingIndex("tcest,k=", source);//なければ0を返す
            if (readingIndex != 0)
            {
                while (source[readingIndex + 1] != '\n')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                inputToCells(BarData.Length + 5, 1, inputDataValue, sheet);
            }
            //酸化剤
            readingIndex = 0;
            inputDataValue = "";
            readingIndex = getReadingIndex("oxid=", source);//なければ0を返す
            if (readingIndex != 0)
            {
                inputToCells(BarData.Length + 4, 4, "酸化剤", sheet);
                getc = source[readingIndex + 1];
                //名前
                while(source[readingIndex + 1] != ' ')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                CFtemp = inputDataValue;
                inputToCells(BarData.Length + 5, 4, inputDataValue, sheet);
                //質量比
                inputDataValue = "";
                readingIndex = getReadingIndex("wt=", source, 1);//1個めのwt=を探す
                getc = source[readingIndex + 1];
                while (source[readingIndex + 1] != ' ')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                inputToCells(BarData.Length + 6, 4, inputDataValue, sheet);
                //温度
                inputDataValue = "";
                readingIndex = getReadingIndex(" t,k=", source, 1);//1個めのt,k=を探す
                getc = source[readingIndex + 1];
                while (source[readingIndex + 1] != '\n')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                inputToCells(BarData.Length + 7, 4, inputDataValue, sheet);
                //追加情報
                if (source[readingIndex + 4] != 'f')//次の行がfuelなら追加情報なし
                {
                    //エネルギーH
                    extraInfoNum += 1;
                    inputDataValue = "";
                    readingIndex = getReadingIndex("h,kj/mol=", source, extraInfoNum);//ExtraInfoNum個めのh,kj/molを探す
                    getc = source[readingIndex + 1];
                    while (source[readingIndex + 1] != ' ')
                    {
                        getc = source[readingIndex + 1];
                        inputDataValue += getc;
                        readingIndex += 1;
                    }
                    inputToCells(BarData.Length + 8, 4, inputDataValue, sheet);
                    //化学式
                    inputDataValue = "";
                    while (source[readingIndex + 1] == ' ')//空白がなくなるまで
                    {
                        getc = source[readingIndex + 1];
                        readingIndex += 1;
                    }
                    while (source[readingIndex + 1] != '\n')//こっから化学式開始
                    {
                        getc = source[readingIndex + 1];
                        if (getc != ' ') //空白じゃなければ追加
                            inputDataValue += getc;
                        readingIndex += 1;
                    }
                    inputToCells(BarData.Length + 9, 4, inputDataValue, sheet);

                }
                else
                {
                    inputToCells(BarData.Length + 9, 4, CFtemp, sheet);
                }
            }

            //燃料
            readingIndex = 0;
            inputDataValue = "";
            readingIndex = getReadingIndex("fuel=", source);//なければ0を返す
            if (readingIndex != 0)
            {
                inputToCells(BarData.Length + 4, 5, "燃料", sheet);
                //名前
                getc = source[readingIndex + 1];
                while (source[readingIndex + 1] != ' ')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                CFtemp = inputDataValue;
                inputToCells(BarData.Length + 5, 5, inputDataValue, sheet);
                //質量比
                inputDataValue = "";
                readingIndex = getReadingIndex("wt=", source, 2);//2個めのwt=を探す
                getc = source[readingIndex + 1];
                while (source[readingIndex + 1] != ' ')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                inputToCells(BarData.Length + 6, 5, inputDataValue, sheet);
                //温度
                inputDataValue = "";
                readingIndex = getReadingIndex(" t,k=", source, 2);//2個めのt,k=を探す
                getc = source[readingIndex + 1];
                while (source[readingIndex + 1] != '\n')
                {
                    getc = source[readingIndex + 1];
                    inputDataValue += getc;
                    readingIndex += 1;
                }
                inputToCells(BarData.Length + 7, 5, inputDataValue, sheet);
                //追加情報
                if (source[readingIndex + 4] != 'f')//次の行がfuelなら追加情報なし
                {
                    //エネルギーH
                    extraInfoNum += 1;
                    inputDataValue = "";
                    readingIndex = getReadingIndex("h,kj/mol=", source, extraInfoNum);//ExtraInfoNum個めのh,kj/molを探す
                    getc = source[readingIndex + 1];
                    while (source[readingIndex + 1] != ' ')
                    {
                        getc = source[readingIndex + 1];
                        inputDataValue += getc;
                        readingIndex += 1;
                    }
                    inputToCells(BarData.Length + 8, 5, inputDataValue, sheet);
                    //化学式
                    inputDataValue = "";
                    while (source[readingIndex + 1] == ' ')//空白がなくなるまで
                    {
                        getc = source[readingIndex + 1];
                        readingIndex += 1;
                    }
                    while (source[readingIndex + 1] != '\n')//こっから化学式開始
                    {
                        getc = source[readingIndex + 1];
                        if (getc != ' ') //空白じゃなければ追加
                            inputDataValue += getc;
                        readingIndex += 1;
                    }
                    inputToCells(BarData.Length + 9, 5, inputDataValue, sheet);

                }
                else
                {
                    inputToCells(BarData.Length + 9, 5, CFtemp, sheet);
                }
            }

            //調整
            sheet.Columns().AdjustToContents();
            sheet.Rows().AdjustToContents();
            var OFBARXrangeStart = sheet.Range(2, 1, 2, BarData.Length + 2);
            OFBARXrangeStart.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            OFBARXrangeStart.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            var OFBARYrangeStart = sheet.Range(1, 2, OFData.Length + 2, 2);
            OFBARYrangeStart.Style.Border.RightBorder = XLBorderStyleValues.Thick;
            OFBARYrangeStart.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            var OFBARXrangeEnd = sheet.Range(OFData.Length + 3, 1, OFData.Length + 3, BarData.Length + 2);
            OFBARXrangeEnd.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            var OFBARYrangeEnd = sheet.Range(1, BarData.Length + 3, OFData.Length + 2, BarData.Length + 3);
            OFBARYrangeEnd.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            var CalcConditionsXrange = sheet.Range(3, BarData.Length + 4, 3, BarData.Length + 9);
            CalcConditionsXrange.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        }

        void calcIndividually(object sender, EventArgs e)//O/F,燃焼室圧力が1つずつ指定された場合に.outからそのC*とγを表示 run by button
        {
            //CSTAR,GAMMA表示
            string readCstarLine = "", readGammaLine = "";
            string[] gammaSplit=null, cstarSplit=null;
            //例外処理
            int Switch = 0;
            //ファイル
            string fileName = this.textBox[19].Text + ".out";
            string filePath = System.IO.Path.Combine("NASA-CEA\\" + fileName);
            System.IO.StreamReader file = null;
            //.out読み込み
            try
            {
                file = new System.IO.StreamReader(filePath);
                Switch = 1;
            }
            catch
            {
                Switch = -1;
            }
            switch (Switch)
            {
                case 1:
                    while ((readGammaLine = file.ReadLine()) != null) 
                    {
                        if (readGammaLine.Contains("GAMMA") == true)
                        {
                            gammaSplit = readGammaLine.Split(' ');
                            break;
                        }
                    }
                    while ((readCstarLine = file.ReadLine()) != null)
                    {
                        if (readCstarLine.Contains("CSTAR") == true)
                        {
                            cstarSplit = readCstarLine.Split(' ');
                            break;
                        }
                    }
                    MessageBox.Show("CSTAR:\t" + cstarSplit[17] + "\n" + "GAMMA:\t" + gammaSplit[16], "CSTAR and GAMMA", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    
                    file.Close();
                    break;

                case -1:
                    MessageBox.Show(fileName + "を開けません", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        void openFileByExcel(object sender, EventArgs e)//excelでファイル開く run by button
        {
            try
            {
                System.Diagnostics.Process.Start("Simulator\\" + this.textBox[19].Text + ".xlsx");
            }
            catch
            {
                MessageBox.Show(this.textBox[19].Text + ".xlsxを開けません。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void radioButtonState(object sender, EventArgs e)//ラジオボタン押下
        {
            if (this.radioButton[0, 0].Checked == true)
                radioButtonManager(1, 1);
            else
                radioButtonManager(1, 2);

            if (this.radioButton[1, 0].Checked == true)
                radioButtonManager(2, 1);
            else
                radioButtonManager(2, 2);
        }

        void radioButtonManager(int row, int column)//ラジオボタンの管理
        {
            if (row == 1)
            {
                if (column == 1)
                {
                    this.item[2].Location = new System.Drawing.Point(350, 52);
                    this.item[2].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[2].Text = "to";
                    this.item[3].Visible = true;
                    this.item[3].Location = new System.Drawing.Point(520, 52);
                    this.item[3].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[3].Text = "間隔";
                    //テキストボックス
                    this.textBox[0].Visible = true;//from
                    this.textBox[1].Visible = true;//to
                    this.textBox[2].Visible = false;//indivisual
                    this.textBox[3].Visible = true;//interval
                    this.textBox[0].Location = new System.Drawing.Point(250, 55);
                    this.textBox[1].Location = new System.Drawing.Point(380, 55);
                    this.textBox[3].Location = new System.Drawing.Point(570, 55);
                }
                else
                {
                    this.item[3].Visible = false;
                    this.textBox[0].Visible = false;
                    this.textBox[1].Visible = false;
                    this.textBox[2].Visible = true;
                    this.textBox[3].Visible = false;
                    this.item[2].Location = new System.Drawing.Point(350, 52);
                    this.item[2].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[2].Text = "「,」で区切り入力";
                    //テキストボックス
                    this.textBox[2].Location = new System.Drawing.Point(235, 55);
                }
            }
            if (row == 2)
            {
                if (column == 1)
                {
                    this.item[5].Location = new System.Drawing.Point(500, 92);
                    this.item[5].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[5].Text = "to";
                    this.item[6].Visible = true;
                    this.item[6].Location = new System.Drawing.Point(670, 92);
                    this.item[6].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[6].Text = "間隔";
                    //テキストボックス
                    this.textBox[4].Visible = true;
                    this.textBox[5].Visible = true;
                    this.textBox[6].Visible = false;
                    this.textBox[7].Visible = true;
                    this.textBox[4].Location = new System.Drawing.Point(400, 95);
                    this.textBox[5].Location = new System.Drawing.Point(530, 95);
                    this.textBox[7].Location = new System.Drawing.Point(720, 95);
                }
                else
                {
                    this.item[6].Visible = false;
                    this.textBox[4].Visible = false;
                    this.textBox[5].Visible = false;
                    this.textBox[6].Visible = true;
                    this.textBox[7].Visible = false;
                    this.item[5].Location = new System.Drawing.Point(510, 92);
                    this.item[5].Font = new System.Drawing.Font("メイリオ", fontSize[0], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
                    this.item[5].Text = "「,」で区切り入力";
                    //テキストボックス
                    this.textBox[6].Location = new System.Drawing.Point(400, 95);
                }
            }
        }

        void readSettings(string settingName, ref string gettingData, string filePath)//defaultSettingsファイル読み込み
        {
            string work = "";
            int getc;
            bool FindLine = false;
            try
            {
                StreamReader file = new StreamReader(filePath);
                while (true)
                {
                    if (file == null) break;
                    getc = file.Read();
                    if ((char)getc != '\n' && (char)getc != '\r')
                        work += (char)getc;
                    if (file.EndOfStream)
                        break;
                    if ((char)getc == '=')
                    {
                        work = work.Remove(work.Length - 1, 1);
                        if (work == settingName)
                        {
                            FindLine = true;
                        }
                        else
                        {
                            file.ReadLine();
                            work = "";
                        }
                    }
                    if (FindLine == true)
                    {
                        gettingData = file.ReadLine();
                        file.Close();
                        break;
                    }
                }
            }
            catch
            {

            }
        }

        void toolTipInitialize(ToolTip obj, Control ctrl, string text)//ツールチップ
        {
            obj.InitialDelay = 1000;
            obj.ReshowDelay = 1000;
            obj.AutoPopDelay = 10000;
            obj.SetToolTip(ctrl, text);
        }
    }

}
