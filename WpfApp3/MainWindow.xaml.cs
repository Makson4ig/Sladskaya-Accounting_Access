using System.Data.OleDb;
using System.Windows;
using System.Data;
using System.Collections.Generic;
using System.Configuration;
using Dapper;
using System.Text.RegularExpressions;
using System.IO;
using System.Data.SqlClient;
using System;
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace WpfApp3
{

    public partial class MainWindow : Window
    {
        OleDbConnection Connect = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = C:\\T.accdb"); // Строка подключения к базе данных Access
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Add_Click(object sender, RoutedEventArgs e) // Добавление предмета 
        {
            Window1 windowsAdd = new Window1();
            windowsAdd.ShowDialog(); // Открывает окно для добавления предмета 
            Refresh(); // Обновляет DataGrid
        }

        private void Refresh()
        {
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("select * from Thing ORder BY id", Connect);  // Переменная заполняющая DataSet данными из БД

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(dataAdapter); // Переменная commandBuilder позволяем автоматически сгенерировать нужные выражения
            DataSet dataSet = new DataSet(); // Переменная хранилища данных 

            dataAdapter.Fill(dataSet, "Thing"); // DataSet для заполнения с записями и Строка, указывающая имя исходной таблицы.
            DateGrid.ItemsSource = dataSet.Tables["Thing"].DefaultView; // Выводим все данные на DataGrid 
            DateGrid.Columns[0].Visibility = Visibility.Hidden; // Скрываем колонку id на DataGrid, она нужна только для вычисления
        }

        private void Delete_Click(object sender, RoutedEventArgs e) // Кнопка удаления задачи из DataGrid и базы Access
        {
            if (DateGrid.SelectedItems != null) // Проверка выбран ли элементы 
            {
                for (int i = 0; i < DateGrid.SelectedItems.Count; i++) // Цикл по выбранным элементам на DataGrid
                {
                    DataRowView datarowView = DateGrid.SelectedItems[i] as DataRowView; // Переменная просмотра строки данных 
                    if (datarowView != null) // Проверка на пустоту переменной
                    {
                        string id = datarowView.Row[0].ToString(); // Переменная id с помощью которой будет удаляться запись из БД 
                        Connect.Open();  // Открываю базу данных
                        DataRow dataRow = (DataRow)datarowView.Row; // Переменаня просмотра строки 
                        Connect.Execute("DELETE * FROM Thing WHERE id = " + (Convert.ToInt16(id))); // Выполнение запроса (Удаление выбранной строки)
                        dataRow.Delete(); // Удаление строки из DataGrid 
                        Connect.Close(); // Закрытие базы данных 
                    }
                }
            }
            else
            {
                MessageBox.Show("База пустая", "Error");
            }
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
           Refresh();
        }

        private void PersonBtn_Click(object sender, RoutedEventArgs e) // Изменение фамилии взявшего предмет
        {
            if(PersonTextBox != null) // Проверяем не пусто PersonTextBox
            {
                if (DateGrid.SelectedItems != null) // Проверяем не пусто выбранный элемент на DataGrid
                {
                    for (int i = 0; i < DateGrid.SelectedItems.Count; i++) // Цикл по выбранным элементам на DataGrid
                    {
                        DataRowView datarowView = DateGrid.SelectedItems[i] as DataRowView; // Переменная просмотра строки данных 
                        if (datarowView != null) // Проверка на пустоту переменной 
                        {
                            string id = datarowView.Row[0].ToString(); // Переменная id с помощью которой будет удаляться запись из БД 
                            Connect.Open(); // Открываю базу данных
                            DataRow dataRow = (DataRow)datarowView.Row;  // Переменаня просмотра строки 
                            Connect.Execute(@"UPDATE Thing SET[Фамилия] = '"+ PersonTextBox.Text + "' where id ="+ Convert.ToInt16(id)); // Обновляем данные в базе 
                            Connect.Close(); // Закрытие базы данных 
                        }
                    }
                }
                else
                {
                    MessageBox.Show("База пустая", "Error");
                }
            }
            Refresh();
        }
    }
}
