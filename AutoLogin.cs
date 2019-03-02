using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

class AutoLogin : Form {
    int form_width;
    int form_height;
    int label_width;
    int label_height;
    int buttons_width;
    int buttons_height;
    Label message_label = new Label();
    Button[] buttons;   // ボタン：要素数は読み込むファイル(setting.csv)により変動
    List<string> data_lists = new List<string>(); //読み込んだファイル(setting.csv)のデータを格納
    string current_dir = System.IO.Directory.GetCurrentDirectory();
    string file_name = "./resource/setting.csv";
    string tool_path_IE = "/resource/autoLogin_IE.vbs";


    public static void Main() {
        Application.Run(new AutoLogin());
    }

    public AutoLogin() {
        // Formのサイズ設定
        form_width = 200;
        form_height = 200;
        label_width = form_width;
        label_height = 30;
        buttons_width = form_width;
        buttons_height = form_height - label_height;
        
        Console.WriteLine("form_width:{0}, form_height:{1}", form_width, form_height);
        Console.WriteLine();
        this.ClientSize = new Size(form_width, form_height);

        // Formのサイズ固定
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
        // Formの最大化無効
        this.MaximizeBox = false;
        this.Text = "AutoLogin";

        GridForm_Layout(this, EventArgs.Empty);
    }

    void GridForm_Layout(object sender, EventArgs e) {        
        StreamReader sr = new StreamReader(file_name, Encoding.GetEncoding("SHIFT_JIS"));
        int i = 1;

        // data_lists[0]は設定ファイル用として確保
        data_lists.Add("設定ファイル");


        while(sr.EndOfStream == false) {
            // 一行読み込み
            string line = sr.ReadLine();
            // 先頭2文字が"//"のものはコメントとして除外
            if(line.Substring(0, 2) != "//") {
                data_lists.Add(line);
                Console.WriteLine("data_Lists[{0}]: {1}", i, data_lists[i]);
                Console.WriteLine();
                i++;
            }   
        }

        sr.Close();
        // ラベルの設定
        message_label.Text = "ボタンを選択してください！";
        // ラベルサイズ
        message_label.SetBounds(0, 0, label_width, label_height);
        // 中央揃え
        message_label.TextAlign = ContentAlignment.MiddleCenter;
        // 太字
        message_label.Font = new Font(this.Font, FontStyle.Bold);
        this.Controls.Add(message_label);

        // ボタン数を格納
        int button_num = data_lists.Count;
        // buttonsの要素数を決定
        buttons = new Button[button_num];
        int findPos;    // 最初の区切り文字が何文字目にあるかを格納
        int row_num = button_num;   // ボタンの行数を格納
        int col_num = 1;    // ボタンの列数を格納
        int cx = buttons_width / col_num;
        int cy = buttons_height / row_num;

        //Console.Write("cx:" + cx + "\t cy:" + cy + "\n");
        // Row：行、Col：列
        int index = 0;
        for(int col = 0; col < col_num; col++) {
            for(int row = 0; row < row_num; row++) {
                if(index < button_num) {
                    // Console.WriteLine("R: " + row +"\t C: " + col);
                    // Console.WriteLine("[" + index + "] " + "s(" + cx * col + ", " + cy * row + ")");
                    
                    // buttonsそれぞれを初期化
                    buttons[index] = new Button();
                    if(index == 0) {
                        buttons[index].Name = "設定ファイル";
                        buttons[index].Text = "設定ファイルを開く";
                    } else {
                        // 最初の区切り文字が何文字目にあるかを格納
                        findPos = data_lists[index].IndexOf(",") + 1;
                        // 0文字目から区切り文字の1文字手前までを抜き出して格納
                        buttons[index].Name = data_lists[index].Substring(0, findPos - 1);
                        buttons[index].Text = data_lists[index].Substring(0, findPos - 1);
                        // ボタンが押されたときに渡す引数を格納
                        buttons[index].Tag = data_lists[index];
                    }

                    // ボタンを設置する位置を決定し配置
                    buttons[index].SetBounds(cx * col, cy * row + label_height, cx, cy);
                    this.Controls.Add(this.buttons[index]);
                    // イベントハンドラを登録
                    buttons[index].Click += new System.EventHandler(button_Click);
                    index++;
                }
            }
        }
    }

    // イベントハンドラ本体
    void button_Click(object sender, EventArgs e) {
        if(((Button)sender).Name == "設定ファイル") {
            MessageBox.Show("設定ファイルOPEN");
            System.Diagnostics.Process p = System.Diagnostics.Process.Start("notepad", file_name);
            // 本アプリの再起動処理

        } else {
            string temp_path = current_dir + tool_path_IE;
            // 押されたボタンのNAMEを表示
            // MessageBox.Show(((Button)sender).Name);

            // 引数の形式：NAME URL ID ID_TYPE PW PW_TYPE SUBMIT
            string args = (((Button)sender).Tag).ToString();
            // MessageBox.Show(args);
            // vbs起動(絶対パスで指定しないとエラー)
            System.Diagnostics.Process p = System.Diagnostics.Process.Start(temp_path, args);

            // 終了するまで最大10秒間だけ待機
            p.WaitForExit(10000);

            // 終了したかを確認
            if(p.HasExited) {
                // MessageBox.Show("終了しました。" + "\n終了コード:" + p.ExitCode.ToString() + "\n終了時間:" + p.ExitTime.ToString());
            } else {
                // MessageBox.Show("タイムアウトしました。");
            }
        }
        
    }

}