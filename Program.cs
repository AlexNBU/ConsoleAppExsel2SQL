using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApp2SQL
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Введите имя группы");
            //Строка соединения читается автоматически из 
            // с генерированого файла конфигурационого  App.config 
            //но через файл с процедурой подключения MyDbContext
            // страница 853
            using (var context = new MyDbContext())
            {







                    var group = new Group();
                    Excel.Application xlApp = new Excel.Application();
                    xlApp.Visible = true;
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\yablonsky.Alexander\Desktop\2019.xlsx");
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    string first = "";
                    string second = "(";
                    string thrid = "(";

                    int count = 0;

                    for (int i = 1; i <= rowCount; i++)
                    {

                        if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                        {

                            first = first + "('" + xlRange.Cells[i, 1].Value2 + "'),";
                            count++;

                        /*
                         //добавить навую строку 
                        group = new Group()
                        {
                            name = xlRange.Cells[i, 1].Value2,
                            name1 = xlRange.Cells[i, 2].Value2,
                            name2 = xlRange.Cells[i, 3].Value2,
                            name3 = xlRange.Cells[i, 4].Value2,
                            Year = count
                        };
                        
                        context.Groups.Add(group);
                        */

                        //context.Groups.Remove
                        context.Groups.Add(new Group()
                        {
                            name = xlRange.Cells[i, 1].Value2,
                            name1 = xlRange.Cells[i, 2].Value2,
                            name2 = xlRange.Cells[i, 3].Value2,
                            name3 = xlRange.Cells[i, 4].Value2,
                            Year = count
                        });
                        //
                        context.SaveChanges();
                        Console.WriteLine($"id: {group.Id}");
                        //Console.Read();



                        continue;

                        }

                    }
                    Console.WriteLine(count);
                    Console.WriteLine();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);

                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    var value1 = first.Remove(first.Length - 1, 1);
                // Console.ReadKey();
                /*
                EntityKey key = new EntityKey(
                    context.Groups,
                    Id,
                    2);
                */
                //Group carToDelete = (Group)context.GetObjectByKey(key);




                //Удилить строку с идентификатором 2 
                if (context.Groups.Find(2) != null)
                    {
                    context.Groups.Remove(context.Groups.Find(2));
                    context.SaveChanges();
                }


/*
                var ff = (from g in context.Groups where g.Id == 3 select g).FirstOrDefault();
                context.Groups.Remove(ff);
                context.SaveChanges();
                */
                //   context.DeleteObject(carToDelete);
                //context.SaveChanges();



                //EntityKey key = new EntityKey(context.group, Id, 2);

                //context.Groups.Add(group);
                context.SaveChanges();
                Console.WriteLine($"id: {group.Id}");
                Console.Read();
                //group = "";
                foreach (string name in context.Groups select name)
                {
                    Console.WriteLine(name);
                }

                Console.WriteLine("jjlkjlkj"); 
                Console.Read();

                //User user1 = new User { Name = "Tom", Age = 33 };
                //User user2 = new User { Name = "Sam", Age = 26 };

                // добавляем их в бд
                //context.Users.Add(user1);
                //context.SaveChanges();
            }
            /*
            using (var context1 = new MyDbContext())
            {
                EntityKey key = new EntityKey(
                    context1.Groups,
                    Id,
                    2);
                Group carToDelete = (Group)context1.GetObjectByKey(key);

                context1.DeleteObject(carToDelete);
                context1.SaveChanges();

            }
            */    


        }
    }
}
