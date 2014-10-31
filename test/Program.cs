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
			
			fxc.Create( "new.xlsx" );
			for ( int i = 0; i < 3; ++i ) {
				fxc.NewWorksheet( "Sheet"+(i+2).ToString() );
			}
			fxc.Worksheets[1].SetRandomCellValuesForRange("A1:W5000", true);
			fxc.Worksheets[0].SetRandomCellValuesForRange("A2:B10", true);

			//	fxc.Worksheets[0].SetTextCellValue("A2", "dwddddwerdwerdwer");
			fxc.Save("new.xlsx");
			
			fxc.Close();
			
			
			fxc.Open("new.xlsx");
			
			
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("B1") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("A2") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("A1") );
			Console.WriteLine( fxc.Worksheets[0].Name+":"+fxc.Worksheets[0].GetCellValue("C57") );
			
			//
			//	fxc.Worksheets[0].SetTextCellValue("A2", "dwddddwerdwerdwer");
			

			
			fxc.Close();

			

			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
		
	}
}