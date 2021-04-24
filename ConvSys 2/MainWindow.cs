using System;
using System.Windows.Forms;

namespace ConvSys_2
{
    public partial class MainWindow : Form
    {
        delegate void Message(string message);
        Message mes;
        public MainWindow()
        {
            InitializeComponent();
            mes = MessageErrorShow;
        }

        private void BTN_BrowseDBFrom_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Excel (*.xlsx)|*.xlsx";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_DBFrom.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseDBnsi_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_DBnsi.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseDBOut_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_DBOut.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_OpenConvertWindow_Click(object sender, EventArgs e)
        {
            if (TB_DBFrom.Text == "")
            {
                mes("Не выбрана конвертируемая таблица Excel");
                return;
            }

            if (TB_DBnsi.Text == "")
            {
                mes("Не выбрана база НСИ");
                return;
            }

            if (TB_DBOut.Text == "")
            {
                mes("Не выбрана итоговая база ЛесИС для конвертирования!");
                return;
            }

            DialogResult dialogResult = MessageBox.Show("Сейчас начнется конвертирование данных. Воизбежании утери нужных данных, сделайте резервную копию итоговой базы ЛесИС. Вы уверины, что хотите продолжить?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if(dialogResult == DialogResult.Yes)
            {
                try
                {
                    ConvertForm convertForm = new ConvertForm(TB_DBFrom.Text, TB_DBnsi.Text, TB_DBOut.Text);
                    convertForm.ShowDialog();
                }
                catch (Exception ex)
                {
                    mes(ex.Message);
                }
            }
            
        }

        private static void MessageErrorShow(string message)
        {
            MessageBox.Show($"{message}","Ошибка", MessageBoxButtons.OK,MessageBoxIcon.Error);
        }
    }
}
