/*
 * Пользователь: Igor.Evdokimov
 * Дата: 31.07.2014
 * Время: 11:39
 */
using System;

namespace test
{
	class Program
	{
		[STAThreadAttribute]
		public static void Main(string[] args)
		{
			Console.WriteLine("FastXcel test!");
			
			
			fastxcel.FastXcel fxc = new fastxcel.FastXcel();
			
			fxc.Open( @"bookxxx.xlsx" );
			//	fxc.Open( @"book.xlsx" );
//			fxc.Worksheets[0].SetTextCellValue("B765", "ЦВАЦУЦВЦУ");
			
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("B1") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("A2") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("A1") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("C57") );

			fxc.Worksheets[0].SetTextCellValue("A2", "dwddddwerdwerdwer");
			fxc.Worksheets[0].SetRandomCellValuesForRange("A3:B10");
			fxc.Save("book1.xlsx");
			
			fxc.Close();

			

			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
		
	}
}