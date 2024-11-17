using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml; // Для работы с Excel
using System.IO; // Для работы с файлами


namespace AutoSalonApp1
{
    public partial class Form1 : Form
    {
        // Коллекция для хранения данных автомобилей
        private List<Car> cars = new List<Car>();

        public Form1()
        {
            InitializeComponent();
            dataGridView1.DataSource = cars;
        }
        // Класс для представления автомобиля
        public class Car
        {
            public int ID { get; set; }
            public string Brand { get; set; }
            public string Color { get; set; }
        }
        // окно для списка
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        // окно для ввода id
        private void TxtID_TextChanged(object sender, EventArgs e)
        {

        }
        // окно для ввода марки 
        private void txtBrand_TextChanged(object sender, EventArgs e)
        {

        }
        // окно для ввода цвета машины 
        private void txtColor_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int id = int.Parse(txtID.Text);  // Поле для ввода ID
            string brand = txtBrand.Text;    // Поле для ввода марки машины
            string color = txtColor.Text;    // Поле для ввода цвета машины

            cars.Add(new Car { ID = id, Brand = brand, Color = color });
            dataGridView1.DataSource = null; // Обновляем привязку данных
            dataGridView1.DataSource = cars;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            int id = int.Parse(txtID.Text);
            Car carToUpdate = cars.Find(c => c.ID == id);

            if (carToUpdate != null)
            {
                carToUpdate.Brand = txtBrand.Text;  // Изменяем марку
                carToUpdate.Color = txtColor.Text;  // Изменяем цвет
                dataGridView1.DataSource = null; // Обновляем привязку данных
                dataGridView1.DataSource = cars;
            }
            else
            {
                MessageBox.Show("Машина с таким ID не найдена.");
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            int id = int.Parse(txtID.Text);
            Car carToDelete = cars.Find(c => c.ID == id);

            if (carToDelete != null)
            {
                cars.Remove(carToDelete);
                dataGridView1.DataSource = null; // Обновляем привязку данных
                dataGridView1.DataSource = cars;
            }
            else
            {
                MessageBox.Show("Машина с таким ID не найдена.");
            }
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                // Устанавливаем лицензию EPPlus для некоммерческого использования
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    // Создаём лист в Excel
                    var ws = excel.Workbook.Worksheets.Add("Автомобили");

                    // Добавляем заголовки колонок
                    ws.Cells[1, 1].Value = "ID";
                    ws.Cells[1, 2].Value = "Марка";
                    ws.Cells[1, 3].Value = "Цвет";

                    // Добавляем данные из списка автомобилей
                    for (int i = 0; i < cars.Count; i++)
                    {
                        ws.Cells[i + 2, 1].Value = cars[i].ID;
                        ws.Cells[i + 2, 2].Value = cars[i].Brand;
                        ws.Cells[i + 2, 3].Value = cars[i].Color;
                    }

                    // Сохраняем файл Excel через диалог сохранения
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Сохраняем Excel-файл на диск
                        FileInfo fi = new FileInfo(saveFileDialog.FileName);
                        excel.SaveAs(fi);

                        // Уведомление об успешном сохранении
                        MessageBox.Show("Данные успешно экспортированы в Excel!");
                    }
                }
            }
            catch (Exception ex)
            {
                // Обрабатываем возможные ошибки
                MessageBox.Show("Ошибка при экспорте данных в Excel: " + ex.Message);
            }
        }

        
    }
}

