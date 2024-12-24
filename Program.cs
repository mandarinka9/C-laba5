using Microsoft.Office.Interop.Excel;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace _5laba_3_
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\Марина\Downloads\LR5-var1.xls";
            var db = new Database(excelFilePath); //созд экземпляра класса

            // Протоколирование
            string logFilePath;
            Console.WriteLine("Введите путь к файлу для протоколирования:");
            logFilePath = Console.ReadLine();
            if (!File.Exists(logFilePath))
            {
                File.Create(logFilePath).Dispose(); //создан и закрыт
            }

            bool running = true;
            while (running)
            {
                Console.WriteLine("Выберите таблицу для взаимодействия:\n1. Движение товаров\n2. Товар\n3. Категория\n4. Магазин\n5. Выход");
                int sheetnum = int.Parse(Console.ReadLine());
                if (sheetnum == 5) { running = false; }
                else
                {
                    Console.WriteLine("Выберите действие:\n1. Просмотр базы данных\n2. Добавление элемента\n3. Корректировка элемента\n4. Удаление элемента\n5. Выполнение четырех запросов\n6. Выход");
                    var choice = Console.ReadLine();
                    switch (choice)
                    {
                        case "1":
                            db.ViewDatabase(logFilePath, sheetnum);
                            break;
                        case "2":
                            switch (sheetnum)
                            {
                                case 1:
                                    db.AddElement1(logFilePath);
                                    break;
                                case 2:
                                    db.AddElement2(logFilePath);
                                    break;
                                case 3:
                                    db.AddElement3(logFilePath);
                                    break;
                                case 4:
                                    db.AddElement4(logFilePath);
                                    break;
                            }

                            break;
                        case "3":
                            switch (sheetnum)
                            {
                                case 1:
                                    db.EditElement1(logFilePath);
                                    break;
                                case 2:
                                    db.EditElement2(logFilePath);
                                    break;
                                case 3:
                                    db.EditElement3(logFilePath);
                                    break;
                                case 4:
                                    db.EditElement4(logFilePath);
                                    break;
                            }
                            break;
                        case "4":
                            db.DeleteElement(logFilePath, sheetnum);
                            break;
                        case "5":
                            db.ExecuteQuery(logFilePath, excelFilePath);
                            break;
                        case "6":
                            running = false;
                            break;
                        default:
                            Console.WriteLine("Некорректный выбор.");
                            break;
                    }
                }
            }
        }
    }

    class Database
    {
        private readonly string _filePath;  //хранит путь к файлу
        private Excel.Application _excelApp; //хр экземпл excel
        private Excel.Workbook _workbook; //хр открытый раб файл excel

        public Database(string filePath)
        {
            _filePath = filePath;
            InitializeExcel();
        }


        private void InitializeExcel()
        {
            _excelApp = new Excel.Application(); 
            _workbook = _excelApp.Workbooks.Open(_filePath); //работаем с данными через СОМ
        }

        public void ViewDatabase(string logFilePath, int sheetnum)
        {
            Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[sheetnum]; //получаем рабочий лист с нужным номером 
            Excel.Range usedRange = movementSheet.UsedRange; //какие ячейки заполнены, диапазон

            var data = Enumerable.Range(1, usedRange.Rows.Count) //последов от 1 до кол-ва строк в usedRange
                .Select(row => Enumerable.Range(1, usedRange.Columns.Count) //созд последов чисел от 1 до количества столбцов
                .Select(col => usedRange.Cells[row, col].Value2).ToArray()).ToList(); //получает зн из каждой ячейки в тек строке, в мас, в список массивов строк

            foreach (var row in data) //проходит по каждому мас строк row в списке data
            {
                Console.WriteLine(string.Join("\t", row)); //объедин эл мас в одну строку
            }

            LogAction(logFilePath, "Просмотр базы данных выполнен.");
        }


        public void AddElement1(string logFilePath)
        {
            try
            {
                // Создаем объект для хранения данных о движении товаров
                var movementData = new MovementData();

                // Запрос данных у пользователя
                Console.WriteLine("Введите данные для добавления в 'Движение товаров':");

                Console.Write("ID операции: ");
                movementData.OperationId = Console.ReadLine();

                Console.Write("Дата (в формате ДД.ММ.ГГГГ): ");
                movementData.Date = DateTime.Parse(Console.ReadLine());

                Console.Write("ID магазина: ");
                movementData.StoreId = Console.ReadLine();

                Console.Write("Артикул: ");
                movementData.ArticleId = Console.ReadLine();

                Console.Write("Тип операции (Поступление/Продажа/Возврат): ");
                movementData.OperationType = Console.ReadLine();

                Console.Write("Количество упаковок, шт: ");
                movementData.PackageCount = Console.ReadLine();

                Console.Write("Наличие карты клиента (да/нет/-): ");
                movementData.ClientCardStatus = Console.ReadLine();

                // Получаем рабочий лист "Движение товаров"
                Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[1];
                Excel.Range usedRange = movementSheet.UsedRange;

                // Определяем следующую пустую строку
                int nextRow = usedRange.Rows.Count + 1;

                // Добавляем данные в новую строку, 
                movementSheet.Cells[nextRow, 1].Value2 = movementData.OperationId;          // ID операции
                movementSheet.Cells[nextRow, 2].Value = movementData.Date;                // Дата
                movementSheet.Cells[nextRow, 3].Value2 = movementData.StoreId;             // ID магазина
                movementSheet.Cells[nextRow, 4].Value2 = movementData.ArticleId;           // Артикул
                movementSheet.Cells[nextRow, 5].Value2 = movementData.OperationType;       // Тип операции
                movementSheet.Cells[nextRow, 6].Value2 = movementData.PackageCount;        // Количество упаковок
                movementSheet.Cells[nextRow, 7].Value2 = movementData.ClientCardStatus;    // Наличие карты клиента

                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Добавлен новый элемент: ID операции {movementData.OperationId}, Дата: {movementData.Date.ToShortDateString()}, ID магазина: {movementData.StoreId}, Артикул: {movementData.ArticleId}, Тип операции: {movementData.OperationType}, Количество упаковок: {movementData.PackageCount}, Наличие карты клиента: {movementData.ClientCardStatus}.");

                Console.WriteLine("Элемент успешно добавлен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при добавлении элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при добавлении элемента. Пожалуйста, проверьте введенные данные.");
            }
        }


        // Класс для хранения данных 
        private class MovementData
        {
            public string OperationId { get; set; }          // ID операции
            public DateTime Date { get; set; }            // Дата
            public string StoreId { get; set; }           // ID магазина
            public string ArticleId { get; set; }             // Артикул
            public string OperationType { get; set; }      // Тип операции
            public string PackageCount { get; set; }          // Количество упаковок
            public string ClientCardStatus { get; set; }   // Наличие карты клиента



            public string Categoryid { get; set; }
            public string Productname { get; set; }
            public string Purchase { get; set; }
            public string Sale { get; set; }
            public int Discount { get; set; }



            public string Categoryname { get; set; }
            public string Agelimit { get; set; }



            public string District { get; set; }
            public string Address { get; set; }


        }
        public void AddElement2(string logFilePath)
        {
            try
            {
                var movementData = new MovementData(); //созд новый объект класса 

                // Запрос данных у пользователя
                Console.WriteLine("Введите данные для добавления в 'Товар':");

                Console.Write("Артикул: ");
                movementData.ArticleId = Console.ReadLine();

                Console.Write("ID категории: ");
                movementData.Categoryid = Console.ReadLine();

                Console.Write("Наименование товара: ");
                movementData.Productname = Console.ReadLine();

                Console.Write("Цена закупки при поступлении, руб: ");
                movementData.Purchase = Console.ReadLine();

                Console.Write("Цена продажи без учёта скидки, руб: ");
                movementData.Sale = Console.ReadLine();

                Console.Write("Скидка при наличии карты клиента, %: ");
                movementData.Discount = int.Parse(Console.ReadLine());




                // Получаем рабочий лист "Товар"
                Excel.Worksheet product = (Excel.Worksheet)_workbook.Sheets[2];
                Excel.Range usedRange = product.UsedRange;

                // Определяем следующую пустую строку
                int nextRow = usedRange.Rows.Count + 1;

                // Добавляем данные в новую строку
                product.Cells[nextRow, 1].Value2 = movementData.ArticleId;          // Артикул
                product.Cells[nextRow, 2].Value2 = movementData.Categoryid;                // ID категории
                product.Cells[nextRow, 3].Value2 = movementData.Productname;            // Наименование товара
                product.Cells[nextRow, 4].Value2 = movementData.Purchase;          // Цена закупки при поступлении, руб
                product.Cells[nextRow, 5].Value2 = movementData.Sale;       // Цена продажи без учёта скидки, руб
                product.Cells[nextRow, 6].Value2 = movementData.Discount;        // Скидка при наличии карты клиента, %


                // Сохраняем изменения в файле
                _workbook.Save();


                LogAction(logFilePath, $"Добавлен новый элемент: Артикул {movementData.ArticleId}, ID категории: {movementData.Categoryid}, Наименование товара: {movementData.Productname}, Цена закупки при поступлении, руб: {movementData.Purchase}, Цена продажи без учёта скидки, руб: {movementData.Sale}, Скидка при наличии карты клиента, %: {movementData.Discount}.");

                Console.WriteLine("Элемент успешно добавлен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при добавлении элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при добавлении элемента. Пожалуйста, проверьте введенные данные.");
            }
        }


        public void AddElement3(string logFilePath)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос данных у пользователя
                Console.WriteLine("Введите данные для добавления в 'Категория':");

                Console.Write("ID категории: ");
                movementData.Categoryid = Console.ReadLine();

                Console.Write("Наименование: ");
                movementData.Categoryname = Console.ReadLine();

                Console.Write("Возрастное ограничение: ");
                movementData.Agelimit = Console.ReadLine();


                // Получаем рабочий лист "Категория"
                Excel.Worksheet categorySheet = (Excel.Worksheet)_workbook.Sheets[3];
                Excel.Range usedRange = categorySheet.UsedRange;

                // Определяем следующую пустую строку
                int nextRow = usedRange.Rows.Count + 1;

                // Добавляем данные в новую строку
                categorySheet.Cells[nextRow, 1].Value2 = movementData.Categoryid;          // ID категории
                categorySheet.Cells[nextRow, 2].Value2 = movementData.Categoryname;                // Наименование
                categorySheet.Cells[nextRow, 3].Value2 = movementData.Agelimit;            // Возрастное ограничение

                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Добавлен новый элемент: ID категории {movementData.Categoryid}, Наименование: {movementData.Categoryname}, Возрастное ограничение: {movementData.Agelimit}.");

                Console.WriteLine("Элемент успешно добавлен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при добавлении элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при добавлении элемента. Пожалуйста, проверьте введенные данные.");
            }
        }

        public void AddElement4(string logFilePath)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос данных у пользователя
                Console.WriteLine("Введите данные для добавления в 'Магазин':");


                Console.Write("ID магазина: ");
                movementData.StoreId = Console.ReadLine();

                Console.Write("Район: ");
                movementData.District = Console.ReadLine();

                Console.Write("Адрес: ");
                movementData.Address = Console.ReadLine();




                // Получаем рабочий лист "Магазин"
                Excel.Worksheet shopSheet = (Excel.Worksheet)_workbook.Sheets[4];
                Excel.Range usedRange = shopSheet.UsedRange;

                // Определяем следующую пустую строку
                int nextRow = usedRange.Rows.Count + 1;

                // Добавляем данные в новую строку
                shopSheet.Cells[nextRow, 1].Value2 = movementData.StoreId;          // ID магазина
                shopSheet.Cells[nextRow, 2].Value2 = movementData.District;                // Район
                shopSheet.Cells[nextRow, 3].Value2 = movementData.Address;            // Адрес

                // Сохраняем изменения в файле
                _workbook.Save();


                LogAction(logFilePath, $"Добавлен новый элемент: ID магазина: {movementData.StoreId}, Район: {movementData.District}, Адрес: {movementData.Address}.");


                Console.WriteLine("Элемент успешно добавлен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при добавлении элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при добавлении элемента. Пожалуйста, проверьте введенные данные.");
            }
        }






        public void EditElement1(string logFilePath)
        {

            var movementData = new MovementData();

            // Запрос ID операции у пользователя для редактирования
            Console.Write("Введите ID операции для редактирования: ");
            movementData.OperationId = Console.ReadLine();

            // Получаем рабочий лист "Движение товаров"
            Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[1];
            Excel.Range usedRange = movementSheet.UsedRange;

            var foundRow = Enumerable.Range(1, usedRange.Rows.Count)
                .Select(row => usedRange.Rows[row]) //возвращает объект строки с номером row
                .FirstOrDefault(row => row.Cells[1, 1].Value?.ToString() == movementData.OperationId);

            if (foundRow == null)
            {
                Console.WriteLine("Элемент с указанным ID операции не найден.");
            }


            // Запрос новых данных у пользователя
            Console.WriteLine("Введите новые данные (оставьте пустым для сохранения старых):");

            // редактирование
            Console.Write($"Дата (текущая: {(movementSheet.Cells[foundRow.Row, 2].Value).ToShortDateString()}): ");
            movementData.Date = DateTime.Parse(Console.ReadLine());
            if (!string.IsNullOrEmpty(movementData.Date.ToShortDateString()))
                movementSheet.Cells[foundRow.Row, 2].Value = movementData.Date;

            Console.Write($"ID магазина (текущий: {movementSheet.Cells[foundRow.Row, 3].Value2}): ");
            movementData.StoreId = Console.ReadLine();
            if (!string.IsNullOrEmpty(movementData.StoreId))
                movementSheet.Cells[foundRow.Row, 3].Value2 = movementData.StoreId;

            Console.Write($"Артикул (текущий: {movementSheet.Cells[foundRow.Row, 4].Value2}): ");
            movementData.ArticleId = Console.ReadLine();
            if (!string.IsNullOrEmpty(movementData.ArticleId))
                movementSheet.Cells[foundRow.Row, 4].Value2 = movementData.ArticleId;

            Console.Write($"Тип операции (текущий: {movementSheet.Cells[foundRow.Row, 5].Value2}): ");
            movementData.OperationType = Console.ReadLine();
            if (!string.IsNullOrEmpty(movementData.OperationType))
                movementSheet.Cells[foundRow.Row, 5].Value2 = movementData.OperationType;

            Console.Write($"Количество упаковок (текущее: {movementSheet.Cells[foundRow.Row, 6].Value2}): ");
            movementData.PackageCount = Console.ReadLine();
            if (!string.IsNullOrEmpty(movementData.PackageCount))
                movementSheet.Cells[foundRow.Row, 6].Value2 = int.Parse(movementData.PackageCount);

            Console.Write($"Наличие карты клиента (текущее: {movementSheet.Cells[foundRow.Row, 7].Value2}): ");
            movementData.ClientCardStatus = Console.ReadLine();
            if (!string.IsNullOrEmpty(movementData.ClientCardStatus))
                movementSheet.Cells[foundRow.Row, 7].Value2 = movementData.ClientCardStatus;


            // Сохраняем изменения в файле
            _workbook.Save();


            LogAction(logFilePath, $"Элемент с ID операции {movementData.OperationId} изменен.");
            Console.WriteLine("Элемент успешно изменен.");

        }

        public void EditElement2(string logFilePath)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос артикула у пользователя для редактирования
                Console.Write("Введите артикул товара для редактирования: ");
                movementData.OperationId = Console.ReadLine();

                // Получаем рабочий лист "Товары"
                Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[2];
                Excel.Range usedRange = movementSheet.UsedRange;

                var foundRow = Enumerable.Range(1, usedRange.Rows.Count)
                    .Select(row => usedRange.Rows[row])
                    .FirstOrDefault(row => row.Cells[1, 1].Value?.ToString() == movementData.OperationId);

                if (foundRow == null)
                {
                    Console.WriteLine("Элемент с указанным артикулом не найден.");
                }


                // Запрос новых данных у пользователя
                Console.WriteLine("Введите новые данные (оставьте пустым для сохранения старых):");

                Console.Write($"ID категории (текущий: {movementSheet.Cells[foundRow.Row, 2].Value2}): ");
                movementData.StoreId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.StoreId))
                    movementSheet.Cells[foundRow.Row, 2].Value2 = movementData.StoreId;

                Console.Write($"Наименование товара (текущий: {movementSheet.Cells[foundRow.Row, 3].Value2}): ");
                movementData.ArticleId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.ArticleId))
                    movementSheet.Cells[foundRow.Row, 3].Value2 = movementData.ArticleId;

                Console.Write($"Цена закупки при поступлении, руб (текущий: {movementSheet.Cells[foundRow.Row, 4].Value2}): ");
                movementData.OperationType = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.OperationType))
                    movementSheet.Cells[foundRow.Row, 4].Value2 = int.Parse(movementData.OperationType);

                Console.Write($"Цена продажи без учёта скидки, руб (текущее: {movementSheet.Cells[foundRow.Row, 5].Value2}): ");
                movementData.PackageCount = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.PackageCount))
                    movementSheet.Cells[foundRow.Row, 5].Value2 = int.Parse(movementData.PackageCount);

                Console.Write($"Скидка при наличии карты клиента, % (текущее: {movementSheet.Cells[foundRow.Row, 6].Value2}): ");
                movementData.ClientCardStatus = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.ClientCardStatus))
                    movementSheet.Cells[foundRow.Row, 6].Value2 = float.Parse(movementData.ClientCardStatus);


                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Элемент с артикулом {movementData.OperationId} изменен.");
                Console.WriteLine("Элемент успешно изменен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при редактировании элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при редактировании элемента. Пожалуйста, проверьте введенные данные.");
            }
        }


        public void EditElement3(string logFilePath)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос ID категории у пользователя для редактирования
                Console.Write("Введите ID категории для редактирования: ");
                movementData.OperationId = Console.ReadLine();

                // Получаем рабочий лист "Категория"
                Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[3];
                Excel.Range usedRange = movementSheet.UsedRange;


                var foundRow = Enumerable.Range(1, usedRange.Rows.Count)
                    .Select(row => usedRange.Rows[row])
                    .FirstOrDefault(row => row.Cells[1, 1].Value?.ToString() == movementData.OperationId);

                if (foundRow == null)
                {
                    Console.WriteLine("Элемент с указанным ID категории не найден.");
                }





                // Запрос новых данных у пользователя
                Console.WriteLine("Введите новые данные (оставьте пустым для сохранения старых):");

                Console.Write($"Наименование (текущий: {movementSheet.Cells[foundRow.Row, 2].Value2}): ");
                movementData.StoreId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.StoreId))
                    movementSheet.Cells[foundRow.Row, 2].Value2 = movementData.StoreId;

                Console.Write($"Возрастное ограничение (текущий: {movementSheet.Cells[foundRow.Row, 3].Value2}): ");
                movementData.ArticleId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.ArticleId))
                    movementSheet.Cells[foundRow.Row, 3].Value2 = movementData.ArticleId;

                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Элемент с ID категории {movementData.OperationId} изменен.");
                Console.WriteLine("Элемент успешно изменен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при редактировании элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при редактировании элемента. Пожалуйста, проверьте введенные данные.");
            }
        }




        public void EditElement4(string logFilePath)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос ID магазина у пользователя для редактирования
                Console.Write("Введите ID магазина для редактирования: ");
                movementData.OperationId = Console.ReadLine();

                // Получаем рабочий лист "Магазины"
                Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[4];
                Excel.Range usedRange = movementSheet.UsedRange;

                var foundRow = Enumerable.Range(1, usedRange.Rows.Count)
                    .Select(row => usedRange.Rows[row])
                    .FirstOrDefault(row => row.Cells[1, 1].Value?.ToString() == movementData.OperationId);

                if (foundRow == null)
                {
                    Console.WriteLine("Элемент с указанным ID магазина не найден.");
                }


                // Запрос новых данных у пользователя
                Console.WriteLine("Введите новые данные (оставьте пустым для сохранения старых):");

                Console.Write($"Район (текущий: {movementSheet.Cells[foundRow.Row, 2].Value2}): ");
                movementData.StoreId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.StoreId))
                    movementSheet.Cells[foundRow.Row, 2].Value2 = movementData.StoreId;


                Console.Write($"Адрес (текущий: {movementSheet.Cells[foundRow.Row, 3].Value2}): ");
                movementData.ArticleId = Console.ReadLine();
                if (!string.IsNullOrEmpty(movementData.ArticleId))
                    movementSheet.Cells[foundRow.Row, 3].Value2 = movementData.ArticleId;



                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Элемент с ID магазина {movementData.StoreId} изменен.");
                Console.WriteLine("Элемент успешно изменен.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при редактировании элемента: {ex.Message}");
                Console.WriteLine("Произошла ошибка при редактировании элемента. Пожалуйста, проверьте введенные данные.");
            }
        }



        public void DeleteElement(string logFilePath, int sheetnum)
        {
            try
            {
                var movementData = new MovementData();

                // Запрос ID операции у пользователя для удаления
                Console.Write("Введите ID для удаления: ");
                movementData.OperationId = Console.ReadLine();

                // Получаем рабочий лист 
                Excel.Worksheet movementSheet = (Excel.Worksheet)_workbook.Sheets[sheetnum];
                Excel.Range usedRange = movementSheet.UsedRange;

                var foundRow = Enumerable.Range(1, usedRange.Rows.Count)
                    .Select(row => usedRange.Rows[row])
                    .FirstOrDefault(row => row.Cells[1, 1].Value?.ToString() == (movementData.OperationId));

                if (foundRow == null)
                {
                    Console.WriteLine("Элемент с указанным ID не найден.");
                }
                Console.WriteLine(foundRow.Row);


                // Удаляем строку
                Excel.Range rowRange = movementSheet.Rows[foundRow.Row];
                rowRange.Delete();


                // Сохраняем изменения в файле
                _workbook.Save();

                LogAction(logFilePath, $"Удален элемент с ID {movementData.OperationId}.");

                Console.WriteLine("Элемент успешно удалён.");
            }
            catch (Exception ex)
            {
                LogAction(logFilePath, $"Ошибка при удалении элемента: {ex.Message}");

                Console.WriteLine("Произошла ошибка при удалении элемента. Проверьте введенные данные.");
            }
        }


        public void ExecuteQuery(string logFilePath, string excelFilePath)
        {
            var db = new Database(excelFilePath); //вщаимод с excel файлом
            string CategoryUnit = "Игрушки на радиоуправлении";
            string ages = "12+";
            string categoryid;
            string street = "Ходунковый";
            List<string> shopsid = new List<string>();
            List<double> tovar = new List<double>();
            double totalCost = 0;
            DateTime Datestart = DateTime.Parse("01/08/2024");
            DateTime Dateend = DateTime.Parse("05/08/2024");

            Console.WriteLine("Определите общую стоимость детских товаров из категории «Игрушки на радиоуправлении 12+»,проданных магазинами Ходункового района за период с 1 по 5 августа включительно.\n ");
            //___________________________________________________________________________________
            Excel.Worksheet categorySheet = (Excel.Worksheet)_workbook.Sheets[3]; // получение листа Категория
            Excel.Range categoryRange = categorySheet.UsedRange;
            var foundRow = Enumerable.Range(1, categoryRange.Rows.Count) //Перебор строк, пока не найдем нужное значение
            .Select(row => categoryRange.Rows[row]) //для пол каждой строки
            .FirstOrDefault(row => row.Cells[1, 2].Value?.ToString() == (CategoryUnit) && row.Cells[1, 3].Value?.ToString() == (ages));

            if (foundRow == null)
            {
                Console.WriteLine("Элемент с указанным ID не найден.");
            }

            categoryid = (categoryRange.Cells[foundRow.Row, 1].Value2).ToString(); //сохр найден строки
            //__________________________________________________________________________________

            categorySheet = (Excel.Worksheet)_workbook.Sheets[4]; // Магазин
            categoryRange = categorySheet.UsedRange;



            for (int k = 2; k <= categoryRange.Rows.Count; k++) //сохраняем все подходящие значения в список, начинаем со 2 строки
            {
                if (categoryRange.Cells[k, 2].value2.ToString() == street) //если зн во 2 столбце соответствует названию района
                {
                    shopsid.Add(categoryRange.Cells[k, 1].value2);
                }
            }
            //__________________________________________________________________________________

            categorySheet = (Excel.Worksheet)_workbook.Sheets[2]; // Товар
            categoryRange = categorySheet.UsedRange;

            for (int k = 2; k <= categoryRange.Rows.Count; k++) //добавляет идентиф товара и его цену, если принадл категории
            {
                if (categoryRange.Cells[k, 2].value2.ToString() == categoryid)
                {
                    tovar.Add(categoryRange.Cells[k, 1].value2);
                    tovar.Add(categoryRange.Cells[k, 5].value2);
                    tovar.Add(categoryRange.Cells[k, 6].value2);

                }
            }
            //__________________________________________________________________________________
            categorySheet = (Excel.Worksheet)_workbook.Sheets[1]; // Движ товаров
            categoryRange = categorySheet.UsedRange;
            for (int k = 2; k < categoryRange.Rows.Count; k++) //по всем строкам
            {
                if (categoryRange.Cells[k, 5].value == "Продажа" && DateTime.Parse((categoryRange.Cells[k, 2].value).ToShortDateString()) >= Datestart && DateTime.Parse((categoryRange.Cells[k, 2].value).ToShortDateString()) <= Dateend)
                {
                    for (int i = 0; i < tovar.Count; i += 3) //
                    {
                        if (tovar[i] == categoryRange.Cells[k, 4].value)
                        {
                            for (int j = 0; j < shopsid.Count; j++)
                            {
                                if (shopsid[j] == categoryRange.Cells[k, 3].value2)
                                {
                                    if (categoryRange.Cells[k, 7].value == "Да")
                                    {

                                        totalCost += categoryRange.Cells[k, 6].value * tovar[i + 1] * (1 - tovar[i + 2]);

                                    }
                                    else
                                    {
                                        totalCost += categoryRange.Cells[k, 6].value * tovar[i + 1];

                                    }
                                }
                            }
                        }
                    }

                }
            }



            Console.WriteLine("Общая стоимость детских товаров: " + totalCost);
        }

        private void LogAction(string logFilePath, string message)
        {
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }

        ~Database()
        {
            _workbook.Close(false);
            _excelApp.Quit();
            Marshal.ReleaseComObject(_workbook);
            Marshal.ReleaseComObject(_excelApp);
        }
    }
}
