using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Web.UI.WebControls;
using Newtonsoft.Json;
using System.Web.Hosting;
using System.Reflection;
using System.Net.Http;

namespace Youth_Rassvet_Test
{
    public partial class Main_Form : Form
    {
        public Main_Form()
        {
            InitializeComponent();
            guna2Panel_Scale.BackColor = System.Drawing.Color.FromArgb(100, 0, 0, 0);
            guna2Button_Answer1.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_Answer2.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_Answer3.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_Answer4.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_Answer5.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_Next.BackColor = System.Drawing.Color.FromArgb(0, 0, 0, 0);
            guna2Button_ScaleS.BackColor = System.Drawing.Color.FromArgb(100, 0, 0, 0);
            guna2Button_ScaleM.BackColor = System.Drawing.Color.FromArgb(100, 0, 0, 0);
            guna2Button_ScaleL.BackColor = System.Drawing.Color.FromArgb(100, 0, 0, 0);
        }  //Загрузка
        System.Drawing.Color Rassvet_DarkBlue = System.Drawing.Color.FromArgb(13, 86, 126);
        System.Drawing.Color Rassvet_Blue = System.Drawing.Color.FromArgb(62, 146, 191);
        System.Drawing.Color Rassvet_LightBlue = System.Drawing.Color.FromArgb(145, 192, 225);
        System.Drawing.Color Rassvet_DarkPeach = System.Drawing.Color.FromArgb(244, 145, 93);
        System.Drawing.Color Rassvet_LigthPeach = System.Drawing.Color.FromArgb(248, 180, 131);
        int Count_Questions = 1;
        int Count_RightAnswers = 0;
        string[] M_Questions = new string[25];
        string[] M_Questions_Sort = new string[25];
        private void guna2Button_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        } //Выход

        private void guna2Button_Scale_Click(object sender, EventArgs e)
        {
            if (guna2Panel_Scale.Visible == false) { guna2Panel_Scale.Visible = true; }
            else { guna2Panel_Scale.Visible = false; }
        } //Разрешение

        private void guna2Button_Start_Click(object sender, EventArgs e)
        {
            Start_Test();
            CheckForPicture();
            guna2Panel_Questions.BackColor = System.Drawing.Color.FromArgb(100, 0, 0, 0);
            guna2HtmlLabel_Questions.Text = "Вопрос " + Count_Questions + ". " + M_Questions_Sort[Count_Questions - 1];
            guna2Panel_Test.Visible = true;
            guna2Panel_Menu.Visible = false;
        } //Начало теста

        private void Start_Test()
        {
            M_Questions[0] = "Что вы будете делать, если человек противоположенных взглядов от ваших начинает наседать на вас и говорить, что ваша позиция \"неверная\"?"; // 1
            M_Questions[1] = "Что вы будете делать, если человек с которым вы ведёте дискуссию, оскорбляет вас и переходит на личности?"; // 1
            M_Questions[2] = "Что вы будете делать, если один или несколько членов отделения будут дискриминировать или унижать человека по каким-либо признакам (например, внешность)?"; // 1
            M_Questions[3] = "Что вы будете делать, если вы заходите в чат, а там идёт конфликт с переходом на личности, обзывательствами и травлей человека за национальность?"; // 1
            M_Questions[4] = "Что вы будете делать, если один из участников будет открыто призывать неподчиняться координатору или руководству Молодёжного Рассвета?"; // 1
            M_Questions[5] = "Что вы будете делать, если в чате отделения один из участников начнет оскорблять других членов отделения / будет ссориться с другими участниками чата?"; // 1
            M_Questions[6] = "Что вы будете делать, если члены отделения начнут отрицать и не признавать правила Молодёжного Рассвета (например протокол безопасности)?"; // 1
            M_Questions[7] = "Что вы будете делать, если вы увидите в чате, что один из участников критикует деятельность Молодёжного или Партии Рассвет?"; // 1
            M_Questions[8] = "Что вы будете делать, если обнаружите где-либо в интернете документы или скриншоты чатов Молодёжного Рассвета в открытом доступе?"; // 2
            M_Questions[9] = "Что вы будете делать, если вы увидите мем или шутку, высмеивающий Молодёжный или Партию Рассвет?"; // 2
            M_Questions[10] = "Что вы будете делать, если узнаете, что один из участников забрал в своё распоряжение канал/аккаунт отделения, и отказывается предоставлять к нему доступ?"; // 2
            M_Questions[11] = "Что вы будете делать, если участник отделения от лица МР будет делать высказывания, которые могут нарушать законы РФ (например, в соцсетях)?"; // 2
            M_Questions[12] = "Что вы будете делать, если вы увидите в комментариях канала вашего отделения сообщения, нарушающие законодательство РФ?"; // 2
            M_Questions[13] = "Что вы будете делать, если в чате кто-то отправил мем/гифку/изображение непристойного или оскорбительного характера?"; // 2
            M_Questions[14] = "Что вы будете делать, если человек будет нецензурно/грубо выражаться в чате отделения или на собрании?"; // 2
            M_Questions[15] = "Что вы будете делать, если про вас обидно, но не оскорбительно пошутил один из участников?"; // 3
            M_Questions[16] = "Что вы будете делать, если вам не нравится то, что пишет или говорит другой участник в чате, но он вас не слушает?"; // 3
            M_Questions[17] = "Что вы будете делать, если участники активно, но безобидно шутят во время рабочего процесса в отделении?"; // 3
            M_Questions[18] = "Что вы будете делать, если вы увидите в чате долгое и активное общение участников на тему, не связанную с деятельностью отделения?"; // 3
            M_Questions[19] = "Что вы будете делать, если член отделения проводит несогласованную агитацию на вступление в другие организации?"; // 3
            M_Questions[20] = "Что вы будете делать, если один из членов отделения начнет без спроса распространять личную информацию других членов отделения?"; // 4
            M_Questions[21] = "Что вы будете делать, если участник Молодежного Рассвета распространяет ложную информацию о членах организации или об организации в целом?"; // 4
            M_Questions[22] = "Что вы будете делать, если участник организации стал запугивать других участников тем, что передаст их данные неким органам, если они не выполнят какие-либо его требования?"; // 4
            M_Questions[23] = "Что вы будете делать, если узнаете об утечке конфиденциальной информации о Молодёжном Рассвете или о его участниках?"; // 4
            M_Questions[24] = "Что вы будете делать, если узнаете, что один из участников связан или связался с некими сторонними органами?"; // 4
            Random random = new Random();
            M_Questions_Sort = M_Questions;
            M_Questions_Sort = M_Questions_Sort.OrderBy(x => random.Next()).ToArray();
        } //Создание вопросов и их сортировка
        private void guna2Button_Answer1_Click(object sender, EventArgs e)
        {
            guna2Button_Answer1.FillColor = Rassvet_DarkBlue;
            guna2Button_Answer2.FillColor = Rassvet_Blue;
            guna2Button_Answer3.FillColor = Rassvet_Blue;
            guna2Button_Answer4.FillColor = Rassvet_Blue;
            guna2Button_Answer5.FillColor = Rassvet_Blue;
            guna2Button_Next.Enabled = true;
        } // Ответ 1
        private void guna2Button_Answer2_Click(object sender, EventArgs e)
        {
            guna2Button_Answer1.FillColor = Rassvet_Blue;
            guna2Button_Answer2.FillColor = Rassvet_DarkBlue;
            guna2Button_Answer3.FillColor = Rassvet_Blue;
            guna2Button_Answer4.FillColor = Rassvet_Blue;
            guna2Button_Answer5.FillColor = Rassvet_Blue;
            guna2Button_Next.Enabled = true;
        } // Ответ 2
        private void guna2Button_Answer3_Click(object sender, EventArgs e)
        {
            guna2Button_Answer1.FillColor = Rassvet_Blue;
            guna2Button_Answer2.FillColor = Rassvet_Blue;
            guna2Button_Answer3.FillColor = Rassvet_DarkBlue;
            guna2Button_Answer4.FillColor = Rassvet_Blue;
            guna2Button_Answer5.FillColor = Rassvet_Blue;
            guna2Button_Next.Enabled = true;
        } // Ответ 3
        private void guna2Button_Answer4_Click(object sender, EventArgs e)
        {
            guna2Button_Answer1.FillColor = Rassvet_Blue;
            guna2Button_Answer2.FillColor = Rassvet_Blue;
            guna2Button_Answer3.FillColor = Rassvet_Blue;
            guna2Button_Answer4.FillColor = Rassvet_DarkBlue;
            guna2Button_Answer5.FillColor = Rassvet_Blue;
            guna2Button_Next.Enabled = true;
        } // Ответ 4
        private void guna2Button_Answer5_Click(object sender, EventArgs e)
        {
            guna2Button_Answer1.FillColor = Rassvet_Blue;
            guna2Button_Answer2.FillColor = Rassvet_Blue;
            guna2Button_Answer3.FillColor = Rassvet_Blue;
            guna2Button_Answer4.FillColor = Rassvet_Blue;
            guna2Button_Answer5.FillColor = Rassvet_DarkBlue;
            guna2Button_Next.Enabled = true;
        } // Ответ 5

        private void guna2Button_Next_Click(object sender, EventArgs e)
        {
            CheckRightAnswer();
            Count_Questions++;
            if (Count_Questions < 15)
            {
                guna2HtmlLabel_Questions.Text = "Вопрос " + Count_Questions + ". " + M_Questions_Sort[Count_Questions - 1];
                guna2Button_Answer1.FillColor = Rassvet_Blue;
                guna2Button_Answer2.FillColor = Rassvet_Blue;
                guna2Button_Answer3.FillColor = Rassvet_Blue;
                guna2Button_Answer4.FillColor = Rassvet_Blue;
                guna2Button_Answer5.FillColor = Rassvet_Blue;
                guna2Button_Next.Enabled = false;
                CheckForPicture();
            }
            if (Count_Questions == 15)
            {
                guna2HtmlLabel_Questions.Text = "Вопрос " + Count_Questions + ". " + M_Questions_Sort[Count_Questions - 1];
                guna2Button_Answer1.FillColor = Rassvet_Blue;
                guna2Button_Answer2.FillColor = Rassvet_Blue;
                guna2Button_Answer3.FillColor = Rassvet_Blue;
                guna2Button_Answer4.FillColor = Rassvet_Blue;
                guna2Button_Answer5.FillColor = Rassvet_Blue;
                guna2Button_Next.Enabled = false;
                guna2Button_Next.Text = "Результат";
                CheckForPicture();
            }
            if (Count_Questions == 16) 
            {
                guna2Panel_Result.Visible = true;
                guna2Panel_Test.Visible = false;
            }
        } //Следущий вопрос
        private void CheckRightAnswer()
        {
            if (M_Questions[Count_Questions - 1] == M_Questions[0])
            {
                if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers +=2; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[1])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[2])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[3])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[4])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[5])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[6])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[7])
            {
                if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[8])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[9])
            {
                if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[10])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers--; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[11])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers--; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[12])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[13])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[14])
            {
                if (guna2Button_Answer3.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[15]) 
            {
                if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[16])
            {
                if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer1.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[17])
            {
                if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[18])
            {
                if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[19])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[20])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[21])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[22])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers--; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[23])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers--; }
            }
            else if (M_Questions[Count_Questions - 1] == M_Questions[24])
            {
                if (guna2Button_Answer2.FillColor == Rassvet_DarkBlue) { Count_RightAnswers += 2; }
                else if (guna2Button_Answer4.FillColor == Rassvet_DarkBlue) { Count_RightAnswers++; }
                else if (guna2Button_Answer5.FillColor == Rassvet_DarkBlue) { Count_RightAnswers--; }
            }
        } //Проверка правильности ответа / Перемешать в соответствии с новым порядком вопросов

        private void CheckForPicture()
        {
            if (M_Questions_Sort[Count_Questions - 1] == M_Questions[0] || M_Questions_Sort[Count_Questions - 1] == M_Questions[1] || M_Questions_Sort[Count_Questions - 1] == M_Questions[2] || M_Questions_Sort[Count_Questions - 1] == M_Questions[3] || M_Questions_Sort[Count_Questions - 1] == M_Questions[4] || M_Questions_Sort[Count_Questions - 1] == M_Questions[5] || M_Questions_Sort[Count_Questions - 1] == M_Questions[6] || M_Questions_Sort[Count_Questions - 1] == M_Questions[7])
            {
                guna2Panel_Test.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.IMG_3208; //Травля
            }
            else if (M_Questions_Sort[Count_Questions - 1] == M_Questions[8] || M_Questions_Sort[Count_Questions - 1] == M_Questions[9] || M_Questions_Sort[Count_Questions - 1] == M_Questions[10] || M_Questions_Sort[Count_Questions - 1] == M_Questions[11] || M_Questions_Sort[Count_Questions - 1] == M_Questions[12] || M_Questions_Sort[Count_Questions - 1] == M_Questions[13] || M_Questions_Sort[Count_Questions - 1] == M_Questions[14])
            {
                guna2Panel_Test.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия316_20250122211749; //Соцсети
            }
            else if (M_Questions_Sort[Count_Questions - 1] == M_Questions[15] || M_Questions_Sort[Count_Questions - 1] == M_Questions[16] || M_Questions_Sort[Count_Questions - 1] == M_Questions[17] || M_Questions_Sort[Count_Questions - 1] == M_Questions[18] || M_Questions_Sort[Count_Questions - 1] == M_Questions[19])
            {
                guna2Panel_Test.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия1_20250317002027; //Бурное обсуждение
            }
            else if (M_Questions_Sort[Count_Questions - 1] == M_Questions[20] || M_Questions_Sort[Count_Questions - 1] == M_Questions[21] || M_Questions_Sort[Count_Questions - 1] == M_Questions[22] || M_Questions_Sort[Count_Questions - 1] == M_Questions[23] || M_Questions_Sort[Count_Questions - 1] == M_Questions[24])
            {
                guna2Panel_Test.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.мр3; //Передача данных
            }

        } //Нужная картинка на фон

        private void EndTest()
        {
            if (Count_RightAnswers < 9)
            {
                guna2Panel_Result.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия364_20250225174542;
            }
            else if (Count_RightAnswers > 8 && Count_RightAnswers < 17)
            {
                guna2Panel_Result.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия363_20250225174154;
            }
            else if (Count_RightAnswers > 16 && Count_RightAnswers < 24)
            {
                guna2Panel_Result.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия363_20250225174213;
            }
            else if (Count_RightAnswers > 23) 
            {
                guna2Panel_Result.BackgroundImage = Youth_Rassvet_Test.Properties.Resources.Без_названия364_20250225174735;
            }
        } //Конечная картинка на фон

        private void guna2Button_Send_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(guna2TextBox_Send.Text) || string.IsNullOrEmpty(Convert.ToString(guna2ComboBox_Send.SelectedItem)))
            {
                MessageBox.Show("Вам нужно заполнить оба поля для отправки", "Внимание");
            }
            else 
            {
                string teg = guna2TextBox_Send.Text;
                string region = RusToEng(guna2ComboBox_Send.Text);
                string answers = Convert.ToString(Count_RightAnswers);
                using (HttpClient client = new HttpClient())
                {
                    try
                    {
                        string url = "https://script.google.com/macros/s/AKfycbxh6tromXp9EJ1GQtSwMvenbefcqZyi_DWfPxC5fbiK195ML-GSdRvbyguPLtsKv1Kj/exec";
                        var content = new StringContent($"Teg={teg}&Region={region}&Answers={answers}", Encoding.UTF8, "application/x-www-form-urlencoded");
                        HttpResponseMessage response = client.PostAsync(url, content).Result;
                        guna2Panel_Send.Visible = false;
                        EndTest();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла ошибка: " + ex.Message);
                    }
                }
            }
        }   // Отправка результатов
        private string RusToEng(string region)
        {
            switch(region)
            {
                case "Алтайский край": { region = "Altai Krai"; break; }
                case "Амурская область": { region = "Amur Oblast"; break; }
                case "Архангельская область": { region = "Arkhangelsk Oblast"; break; }
                case "Астраханская область": { region = "Astrakhan Oblast"; break; }
                case "Белгородская область": region = "Belgorod Oblast"; break;
                case "Брянская область": region = "Bryansk Oblast"; break;
                case "Владимирская область": region = "Vladimir Oblast"; break;
                case "Волгоградская область": region = "Volgograd Oblast"; break;
                case "Вологодская область": region = "Vologda Oblast"; break;
                case "Воронежская область": region = "Voronezh Oblast"; break;
                case "Еврейская автономная область": region = "Jewish Autonomous Oblast"; break;
                case "Забайкальский край": region = "Zabaykalsky Krai"; break;
                case "Ивановская область": region = "Ivanovo Oblast"; break;
                case "Иркутская область": region = "Irkutsk Oblast"; break;
                case "Кабардино-Балкарская Республика": region = "Kabardino-Balkar Republic"; break;
                case "Калининградская область": region = "Kaliningrad Oblast"; break;
                case "Калужская область": region = "Kaluga Oblast"; break;
                case "Камчатский край": region = "Kamchatka Krai"; break;
                case "Карачаево-Черкесская Республика": region = "Karachay-Cherkess Republic"; break;
                case "Кемеровская область": region = "Kemerovo Oblast"; break;
                case "Кировская область": region = "Kirov Oblast"; break;
                case "Костромская область": region = "Kostroma Oblast"; break;
                case "Краснодарский край": region = "Krasnodar Krai"; break;
                case "Красноярский край": region = "Krasnoyarsk Krai"; break;
                case "Курганская область": region = "Kurgan Oblast"; break;
                case "Курская область": region = "Kursk Oblast"; break;
                case "Ленинградская область": region = "Leningrad Oblast"; break;
                case "Липецкая область": region = "Lipetsk Oblast"; break;
                case "Магаданская область": region = "Magadan Oblast"; break;
                case "Москва": region = "Moscow"; break;
                case "Московская область": region = "Moscow Oblast"; break;
                case "Мурманская область": region = "Murmansk Oblast"; break;
                case "Ненецкий автономный округ": region = "Nenets Autonomous Okrug"; break;
                case "Нижегородская область": region = "Nizhny Novgorod Oblast"; break;
                case "Новгородская область": region = "Novgorod Oblast"; break;
                case "Новосибирская область": region = "Novosibirsk Oblast"; break;
                case "Омская область": region = "Omsk Oblast"; break;
                case "Оренбургская область": region = "Orenburg Oblast"; break;
                case "Орловская область": region = "Oryol(Orel) Oblast"; break;
                case "Пензенская область":  region = "Penza Oblast"; break;
                case "Пермский край":  region = "Perm Krai"; break;
                case "Приморский край":  region = "Primorsky Krai"; break;
                case "Псковская область":  region = "Pskov Oblast"; break;
                case "Республика Алтай":  region = "Altai Republic"; break;
                case "Республика Адыгея": region = "Adygey Republic"; break;
                case "Республика Башкортостан":  region = "Bashkortostan Republic"; break;
                case "Республика Бурятия":  region = "Buryatia Republic"; break;
                case "Республика Дагестан": region = "Dagestan Republic"; break;
                case "Республика Ингушетия": region = "ingushetia Republic"; break;
                case "Республика Калмыкия": region = "Kalmik Republic"; break;
                case "Республика Карелия":  region = "Karelia Republic"; break;
                case "Республика Коми":  region = "Komi Republic"; break;
                case "Республика Марий Эл":  region = "Mari El Republic"; break;
                case "Республика Мордовия":  region = "Mordovia Republic"; break;
                case "Республика Саха(Якутия)":  region = "Sakha (Yakutia) Republic"; break;
                case "Республика Северная Осетия": region = "North Osetia Republic"; break;
                case "Республика Татарстан":  region = "Tatarstan Republic"; break;
                case "Республика Тыва":  region = "Tuva Republic"; break;
                case "Республика Хакасия":  region = "Khakassia Republic"; break;
                case "Ростовская область":  region = "Rostov Oblast"; break;
                case "Рязанская область":  region = "Ryazan Oblast"; break;
                case "Самарская область":  region = "Samara Oblast"; break;
                case "Санкт-Петербург":  region = "Saint Petersburg"; break;
                case "Саратовская область":  region = "Saratov Oblast"; break;
                case "Сахалинская область":  region = "Sakhalin Region"; break;
                case "Свердловская область":  region = "Sverdlovsk Oblast"; break;
                case "Смоленская область":  region = "Smolensk Oblast"; break;
                case "Ставропольский край": region = "Stavrapol Krai"; break;
                case "Тамбовская область": region = "Tambov Oblast"; break;
                case "Тверская область": region = "Tver Oblast"; break;
                case "Томская область": region = "Tomsk Oblast"; break;
                case "Тульская область": region = "Tula Oblast"; break;
                case "Тюменская область": region = "Tyumen Oblast"; break;
                case "Удмуртская Республика": region = "Udmurt Republic"; break;
                case "Ульяновская область": region = "Ulyanovsk Oblast"; break;
                case "Хабаровский край": region = "Khabarovsk Kray"; break;
                case "Ханты-Мансийский автономный округ - Югра": region = "Khanty-Mansiysk Autonomous Okrug - Yugra"; break;
                case "Челябинская область": region = "Chelyabinsk Oblast"; break;
                case "Чеченская Республика": region = "Chechen Republic"; break;
                case "Чувашская Республика": region = "Chuvash Republic"; break;
                case "Чукотский автономный округ": region = "Chukotka Autonomous Oblast"; break;
                case "Ямало-Ненецкий автономный округ": region = "Yamalo-Nenets Autonomous Okrug "; break;
                case "Ярославская область": region = "Yaroslavl Oblast"; break;
            }
            return region;
        }
        private void guna2Button_Screenshot_Click(object sender, EventArgs e)
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd != IntPtr.Zero)
            {
                RECT rect;
                GetWindowRect(hwnd, out rect);
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                using (Bitmap bmp = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    { g.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy); }
                    Clipboard.SetImage(bmp);
                    MessageBox.Show("Скриншот скопирован в буфер обмена.");
                }
            }
            else { MessageBox.Show("Не удалось сделать скриншот окна."); }
        } // Скриншот
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private void guna2Button_Note_Click(object sender, EventArgs e)
        {
            guna2Panel_Note.Visible = false;
        } // Закрыть примечание

        private void guna2Button_ScaleS_Click(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.WindowState = FormWindowState.Normal;
            Form.ActiveForm.Size = new Size(854, 480);
            this.CenterToScreen();
        }  // Разрешение 854х480
        private void guna2Button_ScaleM_Click(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.WindowState = FormWindowState.Normal;
            Form.ActiveForm.Size = new Size(1280, 720);
            this.CenterToScreen();

        }  // Разрешение 1280х720
        private void guna2Button_ScaleL_Click(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.WindowState = FormWindowState.Normal;
            Form.ActiveForm.Size = new Size(1920, 1080);
            this.CenterToScreen();
        }  // Разрешение 1920х1080
        private void guna2Button_ScaleAll_Click(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
        }  // Разрешение На весь экран
    }
}
