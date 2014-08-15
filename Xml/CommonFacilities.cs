using System;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.IO;


using System.Security.Cryptography;

namespace CommonFacilities
{
	public class SendMailException : System.Exception
	{
		public SendMailException(string message) : base(message)
		{
		}
	}

	/// <summary>
	/// Функции для общего пользования
	/// </summary>
	public class Common
	{
		/// <summary>
		/// Открываем файл в ассоциированном приложении
		/// </summary>
		/// <param name="filename"></param>
		static public void OpenInApp(string filename)
		{
			// Открываем в ассоциированном приложении
			System.Diagnostics.Process prc = new System.Diagnostics.Process();
			prc.StartInfo.FileName = filename;
			prc.Start();
		}

		/// <summary>
		/// Число ли введено в строке?
		/// </summary>
		/// <param name="text"></param>
		/// <returns></returns>
		static public bool IsNumber(string text)
		{
			System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^[-+]?[0-9]*\.?[0-9]+$");
			return regex.IsMatch(text);
		}
		
		/// <summary>
		///  Получаем маску подсети
		/// </summary>
		/// <param name="bits"></param>
		static public string GetSubnetMask( string bits ) {
			string ip_addr = "255.255.255.";
			int i_bits =  int.Parse(bits.Trim());
			if ( i_bits >= 24 ) {
				ip_addr += (255 - (int)(Math.Pow(2,32-i_bits))+1).ToString();
			}
			return ip_addr;
		}

		/// <summary>
		/// Разделение строки по разделителю
		/// </summary>
		/// <param name="source">исходная трочка</param>
		/// <param name="delimiter">разделитель</param>
		/// <returns></returns>
		static public List<string> ExplodeString(string source, string delimiter)
		{
			List<string> strgs = new List<string>();
			int i = -1;

			for (; ; )
			{
				i = source.IndexOf(delimiter);
				if (i > -1)
				{
					strgs.Add(source.Substring(0, i));
					source = source.Remove(0, i + delimiter.Length);
				}
				else
				{
					strgs.Add(source);
					break;
				};
			}
			return strgs;
		}

		static public List<string> ExplodeString(string source, string delimiter, string quote)
		{
			List<string> strgs = ExplodeString(source, delimiter);
			for (int i = 0; i < strgs.Count; ++i)
			{
				strgs[i] = quote + strgs[i] + quote;
			}

			return strgs;
		}

		/// <summary>
		/// list<> => string
		/// </summary>
		/// <param name="list">исходный массив List<></param>
		/// <param name="delimiter">разделитель</param>
		/// <returns></returns>
		static public string ListToString(List<string> list, string delimiter)
		{
			string strg = "";
			for (int i = 0; i < list.Count; ++i)
				if (i < list.Count - 1)
					strg += list[i] + delimiter;
				else
					strg += list[i];
			return strg;
		}

		/// <summary>
		/// list<> => string
		/// </summary>
		/// <param name="list">исходный массив List<></param>
		/// <param name="delimiter">разделитель</param>
		/// <param name="quote">кавычка</param>
		/// <returns></returns>
		static public string ListToString(List<string> list, string delimiter, string quote)
		{
			string strg = "";
			for (int i = 0; i < list.Count; ++i)
				if (i < list.Count - 1)
					strg += quote + list[i] + quote + delimiter;
				else
					strg += quote + list[i] + quote;
			return strg;
		}

		/// <summary>
		/// Смена кодировки
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <param name="dest_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncoding(String strg, String src_encoding, string dest_encoding)
		{
			// перекодируем в принятую на сервере кодировку
			return Encoding.GetEncoding(dest_encoding).GetString(Encoding.GetEncoding(src_encoding).GetBytes(strg.ToCharArray(), 0, strg.Length));
		}

		/// <summary>
		/// Смена кодировки
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <param name="dest_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncoding(String strg, String src_encoding, int dest_encoding)
		{
			// перекодируем в принятую на сервере кодировку
			return Encoding.GetEncoding(dest_encoding).GetString(Encoding.Convert(Encoding.GetEncoding(src_encoding), Encoding.GetEncoding(dest_encoding), Encoding.GetEncoding(src_encoding).GetBytes(strg)));
		}

		/// <summary>
		/// Смена кодировки
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <param name="dest_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncoding(String strg, int src_encoding, int dest_encoding)
		{
			// перекодируем в принятую на сервере кодировку
			return Encoding.GetEncoding(dest_encoding).GetString(Encoding.Convert(Encoding.GetEncoding(src_encoding), Encoding.GetEncoding(dest_encoding), Encoding.GetEncoding(src_encoding).GetBytes(strg)));

		}

		/// <summary>
		/// Смена кодировки
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <param name="dest_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncoding(String strg, int src_encoding, string dest_encoding)
		{
			// перекодируем в принятую на сервере кодировку
			return Encoding.GetEncoding(dest_encoding).GetString(Encoding.Convert(Encoding.GetEncoding(src_encoding), Encoding.GetEncoding(dest_encoding), Encoding.GetEncoding(src_encoding).GetBytes(strg)));
		}

		/// <summary>
		/// Смена кодировки
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <param name="dest_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncoding(String strg, Encoding src_encoding, Encoding dest_encoding)
		{
			// перекодируем в принятую на сервере кодировку
			return dest_encoding.GetString(Encoding.Convert(src_encoding, dest_encoding, src_encoding.GetBytes(strg)));
		}


		/// <summary>
		/// Смена кодировки на кодировку по умолчанию
		/// </summary>
		/// <param name="strg"></param>
		/// <param name="src_encoding"></param>
		/// <returns></returns>
		static public String ChangeEncodingToDefault(String strg, Encoding src_encoding)
		{
			return ChangeEncoding(strg, src_encoding, Encoding.Default);
		}

		static public String ChangeEncodingToDefault(String strg, String src_encoding)
		{
			return ChangeEncoding(strg, Encoding.GetEncoding(src_encoding), Encoding.Default);
		}
	}
}