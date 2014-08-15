using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CommonObjects
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

		
		static public string DictionaryToString(Dictionary<string, string> dict,
		                                        string delimiter = "=",
		                                        string key_quote = "",
		                                        string value_quote = "\"",
		                                        string string_delimeter = ",") 
		{
			string strg = "";
			uint i = 0;
			foreach ( KeyValuePair<string, string> kvp in dict ) {
				strg += key_quote + kvp.Key + key_quote + " " + delimiter + " " + value_quote + kvp.Value + value_quote;
				if ( i != dict.Count-1 )
					strg += " " + string_delimeter +"\r\n";
							
				++i;
			}
			
			return strg;
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
		
		public class ParameterDictionary<T_key, T_val>
		{
			public ParameterDictionary() {
				inner_dictionary = new Dictionary<T_key, T_val>();
			}
			
			
			public T_val this[T_key key] {
				get {
					return inner_dictionary[key];
				} set {
					inner_dictionary[key] = value;
				}
			}
			
			public IEnumerator<KeyValuePair<T_key, T_val>> GetEnumerator() {
				foreach ( KeyValuePair<T_key, T_val> kvp in inner_dictionary )
					yield return kvp;
			}
			
			private Dictionary<T_key, T_val> inner_dictionary;
			
		}

	}
}