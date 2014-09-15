/*
 * Пользователь: Igor.Evdokimov
 * Дата: 31.07.2014
 * Время: 11:52
 */
using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;

namespace fastxcel
{
	/// <summary>
	/// Description of Worksheet.
	/// </summary>
	public class Worksheet
	{
		public Dictionary<KeyValuePair<XmlNode, string>, KeyValuePair<string,string>> Contents { get; private set; }
		KeyValuePair<KeyValuePair<XmlNode, string>, KeyValuePair<string,string>>[] tmp_cont_array;
		
		public string Name { get; private set; }
		public string XmlPath { get; private set; }
		public uint Number { get;  set; }
		public string ResourceId { get; private  set; }
		public short SheetId { get; private  set; }
		private AmusingXml.XmlFun xml_fun;
		private Dictionary<int, string> shared_strings;
		XmlNode sheet_data = null;
		
		public Worksheet(Dictionary<int, string> _shared_strings)
		{
			XmlPath = String.Empty;
			Number = 1;
			Name = "Sheet1";
			SheetId = -1;
			Contents = new Dictionary<KeyValuePair<XmlNode, string>, KeyValuePair<string, string>>();
			xml_fun = new AmusingXml.XmlFun();
			shared_strings = _shared_strings;
		}
		
		public void Load( string xml_path, Dictionary<int, string> shared_strings )	{
			XmlPath = xml_path;
			xml_fun.LoadDocument( xml_path );
			LoadData();
		}
		
		public void Load( XmlDocument xml_sheet, Dictionary<int, string> shared_strings )
		{
			xml_fun.LoadDocument( xml_sheet );
			LoadData();
		}
		
		public void Load( string _resource_id, short _sheet_id, string name, string xml_path, Dictionary<int, string> shared_strings )
		{
			XmlPath = xml_path;
			xml_fun.LoadDocument( xml_path );
			LoadData();
			SheetId = _sheet_id;
			ResourceId = _resource_id;
			Name = name;

			Console.WriteLine("fastxcel: worksheet# "+SheetId.ToString()+", name="+Name);
		}
		
		public void Save( string xml_file_path ) {
			// Sort...
			xml_fun.SortNodesByAttrIntValue("row", "r");
			xml_fun.SortNodesByAttrValue("c", "r");
			// Save...
			xml_fun.SaveDocument();
		}
		
		public void Save() {
			// Sort...
			xml_fun.SortNodesByAttrIntValue("row", "r");
			xml_fun.SortNodesByAttrValue("c", "r");
			// Save...
			xml_fun.SaveDocument(XmlPath);
		}
		
		/// <summary>
		/// Extracting row numebr from cell
		/// Example: A12 => 12; BC923 => 923
		/// </summary>
		/// <param name="cell_addr"></param>
		private int ExtractRowNumber(string cell_addr) {
			int ret = -1;
			string strg = cell_addr;
			while ( strg != String.Empty ) {
				if ( !CommonFacilities.Common.IsNumber(strg[0].ToString()) ) {
					strg=strg.Remove(0,1);
				}
				
				if (CommonFacilities.Common.IsNumber(strg) == true)
					break;
			}
			if ( strg != string.Empty )
				ret = int.Parse(strg);
			return ret;
		}
		
		private void SetExistingCellValue( string cell, string new_value ) {
			Dictionary<string, string> search_attrs = new Dictionary<string, string>();
			search_attrs["r"]=cell;
			XmlElement fnd_elem = xml_fun.FindFirstNode( "c", search_attrs );
			if ( fnd_elem != null ) {
				fnd_elem.InnerXml = "<v>"+new_value+"</v>";
			}
			xml_fun.FillXmlElements( xml_fun.Document );
		}
		
		private int FindNumberOfSharedString(string _value) {
			foreach ( var key_vl in shared_strings ) {
				if ( key_vl.Value == _value )
					return key_vl.Key;
			}
			return -1;
		}
		
		private void SetTextContentsValue( string _search_term, string _value ) {
			bool have_it = false;

			// Looking for existant cells
			for (int i = 0; i < tmp_cont_array.Length; ++i ) {
				//Contents.
				var cont = tmp_cont_array[i];

				if ( cont.Key.Value == _search_term ) {
					Contents.Remove(cont.Key); // deleting old values

					int new_sh_string = GenNewSharedString(_value);
					
					cont = new KeyValuePair<KeyValuePair<XmlNode, string>, KeyValuePair<string, string>>(
						cont.Key, new KeyValuePair<string,string>(cont.Value.Key,  new_sh_string.ToString())
					);
					
					have_it = true;
					Contents.Add(cont.Key, cont.Value); // Adding updated value
					SetExistingCellValue(_search_term, new_sh_string.ToString());
					
					break;
				}
			}
			
			// If we haven't found our cell => create new
			if ( !have_it ) {
				XmlNode row_node = null;
				int row_number = ExtractRowNumber( _search_term );
				
				// check if this row already exists
				Dictionary<string,string> chk_rows_attrs = new Dictionary<string, string>();
				chk_rows_attrs["r"]= row_number.ToString();
				XmlElement chk_row = xml_fun.FindFirstNode( "row", chk_rows_attrs );
				
				// If nothing was found - let's create new row
				if ( chk_row == null  ) {
					row_node = xml_fun.Document.CreateNode( XmlNodeType.Element, "", "row", null);
					row_node.Attributes.Append(xml_fun.Document.CreateAttribute("r"));
					row_node.Attributes["r"].Value = row_number.ToString(); // Row number
				} else
					row_node = chk_row;
				
				XmlNode cell_node = xml_fun.Document.CreateNode( XmlNodeType.Element, "", "c", null);
				cell_node.Attributes.Append(xml_fun.Document.CreateAttribute("r"));
				cell_node.Attributes.Append(xml_fun.Document.CreateAttribute("t"));
				cell_node.Attributes["r"].Value =  _search_term;
				cell_node.Attributes["t"].Value =  "s"; // shared_String
				cell_node.InnerXml = cell_node.InnerXml.Replace("xmlns=\"\"", String.Empty); // xmlns=... необходимо удалить

				int new_sh_string = GenNewSharedString(_value);
				
				cell_node.InnerXml = "<v>" + new_sh_string.ToString() + "</v>";
				row_node.AppendChild( cell_node );
				
				row_node.InnerXml = row_node.InnerXml.Replace("xmlns=\"\"", String.Empty); // xmlns=... необходимо удалить
				sheet_data.AppendChild( row_node );
				sheet_data.InnerXml =  sheet_data.InnerXml.Replace("xmlns=\"\"", String.Empty); // xmlns=... необходимо удалить
				
				KeyValuePair<XmlNode, string> key_part = new KeyValuePair<XmlNode, string>(cell_node, _search_term);
				KeyValuePair<string, string> value_part = new KeyValuePair<string, string>(String.Empty, _value);
				
				xml_fun.FillXmlElements( xml_fun.Document );
				
				Contents.Add( key_part, value_part );
			}
			
			tmp_cont_array = Contents.ToArray();
		}
		
		
		int GenNewSharedString(string _value)
		{
			int new_sh_string = FindNumberOfSharedString(_value);
			if (new_sh_string < 0) {
				new_sh_string = GetNewSharedStringNumber();
				shared_strings.Add(new_sh_string, _value);
			}
			return new_sh_string;
		}

		public int GetNewSharedStringNumber() {
			int max_key = -1;
			foreach ( KeyValuePair<int, string> sh_str in shared_strings ) {
				if ( sh_str.Key > max_key )
					max_key = sh_str.Key;
			}
			return max_key+1;
		}
		
		private string GetContentsValue( string _search_term ) {
			string ret = null;
			foreach ( var cont in Contents ) {
				if ( cont.Key.Value == _search_term ) {
					if ( cont.Value.Key == "s" ) {
						int shared_string_nbr = int.Parse(cont.Value.Value);
						if (  shared_strings.ContainsKey( shared_string_nbr ) ) {
							ret = shared_strings[shared_string_nbr];
						}
					} else
						ret = cont.Value.Value;
					break;
				}
			}
			return ret;
		}
		
		public void SetDimensions( string dims ) {
			XmlElement dim = xml_fun.FindFirstNode( "dimension", new Dictionary<string, string>() );
			if ( dim != null ) {
				dim.Attributes["ref"].Value = dims;
				xml_fun.FillXmlElements(xml_fun.Document);
			}
		}
		
		public string GetDimensions() {
			XmlElement dim = xml_fun.FindFirstNode( "dimension", new Dictionary<string, string>() );
			if ( dim != null ) {
				return dim.Attributes["ref"].Value.ToString();
			}
			return String.Empty;
		}
		
		public string GetCellValue( string cell_name ) {
			return GetContentsValue( cell_name );
		}
		
		public void SetTextCellValue( string cell_name, string cell_value ) {
			SetTextContentsValue( cell_name, cell_value );
		}
		
		public void SetTextCellValueForRange(string cells_range, string cells_value ) {
			
			List<string> cells = GetCellsInRange( cells_range );
			int cnt_2 = (int)(Math.Truncate((double)(cells.Count()/2)));
			int full_cnt = cells.Count();
			for ( int i = 0; i <= cnt_2; ++i )  {
				SetTextContentsValue( cells[i], cells_value );
				SetTextContentsValue( cells[full_cnt-i-1], cells_value );
			}
		}
		
		public void SetRandomCellValuesForRange(string cells_range ) {
			Random rand = new Random( DateTime.Now.Millisecond*DateTime.Now.Minute );

			List<string> cells = GetCellsInRange( cells_range );
			int cnt_2 = (int)(Math.Truncate((double)(cells.Count()/2)));
			int full_cnt = cells.Count();
			for ( int i = 0; i <= cnt_2; ++i )  {
				SetTextContentsValue( cells[i], rand.Next().ToString() );
				SetTextContentsValue( cells[full_cnt-i-1], rand.Next().ToString() );
			}
		}
		
		/// <summary>
		/// Selecting cell names in range.
		/// Example: range A1:B10, cell_names: A1,A2...B1,B2...B10
		/// </summary>
		/// <param name="range"></param>
		/// <returns></returns>
		private List<string> GetCellsInRange( string range ) {
			string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			int alphabet_count = alphabet.Count();
			int excel_max_rows = 65535;
			List<string> ret = new List<string>();
			
			System.Text.RegularExpressions.Regex range_regex = new System.Text.RegularExpressions.Regex("([A-Z|^[0-9]]*)([0-9]*):([A-Z|^[0-9]]*)([0-9]*)");
			
			System.Text.RegularExpressions.MatchCollection mtc = range_regex.Matches(range);

			if (mtc.Count < 1)
				throw new FastXcelException( "Error in parsing cells range!" );
			
			if (mtc[0].Groups.Count < 5)
				throw new FastXcelException( "Error in parsing cells range!" );
			
			string range_first_letter = mtc[0].Groups[1].Value.ToString(),
			range_last_letter =  mtc[0].Groups[3].Value.ToString();
			int range_first_nbr =  int.Parse(mtc[0].Groups[2].Value.ToString()),
			range_last_nbr = int.Parse(mtc[0].Groups[4].Value.ToString());
			
			string curr_ltr = range_first_letter;
			int curr_nbr = range_first_nbr;
			int i = 0;
			while ( curr_ltr != range_last_letter && curr_nbr != range_last_nbr ) {
				int tmp = 0;
				curr_ltr = String.Empty;
				while ( tmp < Math.Truncate((double)(i /alphabet_count)) ) {
					curr_ltr += alphabet[i%alphabet_count+alphabet.IndexOf(range_first_letter[0])];
					++tmp;
				}
				for ( int k = range_first_nbr; k <= range_last_nbr && k <= excel_max_rows; ++k  ) {
					if ( curr_ltr != String.Empty )
						ret.Add( curr_ltr + k.ToString());
				}
				++i;
			}
			
			return ret;
		}
		
		private void LoadData()
		{
			List<XmlNode> cells = new List<XmlNode>();

			foreach (XmlElement el in xml_fun.XmlElements) {
				if (el.Name == "sheetData") {
					sheet_data = el;
					//sheet_data.InnerXml = el.InnerXml;
					break;
				}
			}

			if (sheet_data != null) {
				foreach (XmlNode el in sheet_data.ChildNodes) {
					if (el.Name == "row") {
						foreach (XmlNode cell_node in el.ChildNodes) {
							if (cell_node.Name == "c") {
								foreach (XmlNode value_node in cell_node.ChildNodes) {
									if (value_node.Name == "v") {
										if ( cell_node.Attributes["t"] != null ) {
											if ( cell_node.Attributes["t"].Value == "s" ) {
												Contents.Add(
													new KeyValuePair<XmlNode, string>(cell_node,cell_node.Attributes["r"].Value),
													new KeyValuePair<string, string>("s", value_node.InnerText));
											}
										} else {
											Contents.Add(
												new KeyValuePair<XmlNode, string>(cell_node,cell_node.Attributes["r"].Value),
												new KeyValuePair<string, string>(String.Empty, value_node.InnerText));
										}
									}
								}
							}
						}
					}
				}
				
				tmp_cont_array = Contents.ToArray();
				
			}
		}
	}
}
