/*
 * Пользователь: Igor.Evdokimov
 * Дата: 31.07.2014
 * Время: 10:42
 */
using System;
using System.IO;
using System.Collections.Generic;
using System.IO.Compression;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;
using AmusingXml;
using System.Xml;

namespace fastxcel
{
	/// <summary>
	/// Description of MyClass.
	/// </summary>
	public class FastXcel : IDisposable
	{
		List<XmlFun> xml_sheets = new List<XmlFun>();
		XmlFun xml_shared_strings;
		XmlFun xml_workbook;
		XmlFun xml_wb_rels;
		Dictionary<int, string> shared_strings = new Dictionary<int, string>();
		FastZip fast_zip;
		string temp_path;
		public List<Worksheet> Worksheets { get; private set; }
		
		public FastXcel() {
			Load();
		}

		void Load() {
			fast_zip = new FastZip();
			xml_shared_strings = new XmlFun();
			xml_sheets = new List<XmlFun>();
			xml_workbook = new XmlFun();
			xml_wb_rels = new XmlFun();
			shared_strings = new Dictionary<int, string>();
			Worksheets = new List<Worksheet>();
		}
		
		void LoadDocPrepare() {
			temp_path = Path.GetTempPath() + Guid.NewGuid().ToString() + "\\";
			Directory.CreateDirectory(temp_path);
		}
		
		private void Extract( string file_path ) {
			try {
				Console.WriteLine("fastxcel: extracting "+file_path+". Temp dir: "+temp_path );
				fast_zip.ExtractZip( file_path, temp_path, "" );
			} catch ( Exception ex ) {
				throw new FastXcelException(ex.Message, ex);
			}
		}
		
		private void Compress( string file_path ) {
			try {
				Console.WriteLine("fastxcel: compressing "+file_path+". Temp dir: "+temp_path );
				fast_zip.CreateZip( file_path, temp_path, true,  "",  "" );
			} catch ( Exception ex ) {
				throw new FastXcelException(ex.Message, ex);
			}
		}
		
		public void Open( string file_path ) {
			try {
				Load();
				LoadDocPrepare();
				Console.WriteLine("fastxcel: opening "+file_path);
				// UnZipping...
				Extract(file_path);
				
				// Process document relationships
				ProcessDocumentRelationships();
				
				// Process workbook main info...
				ProcessWorkbook();
				
				// Reading Shared strings into an array...
				ReadSharedStrings();
				
				// Getting worksheets...
				ProcessWorksheets();
			} catch ( Exception ex ) {
				throw new FastXcelException(ex.Message, ex);
			}
		}
		
		
		public void Create( string file_path ) {
			try {
				Console.WriteLine("fastxcel: creating "+file_path);
				
				System.Resources.ResourceManager resourceManager =
					new System.Resources.ResourceManager("fastxcel.Resource", GetType().Assembly);
				
				
				var assembly = System.Reflection.Assembly.GetExecutingAssembly();
				
				string resource_path = assembly.GetName().Name+".Resources.template_xlsx";
				
				BinaryReader br = new BinaryReader(assembly.GetManifestResourceStream(resource_path));
				long file_length = assembly.GetManifestResourceStream(resource_path).Length;
				byte[] inners = new byte[file_length];
				inners = br.ReadBytes((int)file_length);
				
				br.Close();

				FileStream fs = File.Create( file_path );
				BinaryWriter bw = new BinaryWriter( fs );
				
				bw.Write(inners);
				bw.Close();
				
				Open( file_path );
				
			} catch ( Exception ex ) {
				throw new FastXcelException(ex.Message, ex);
			}
		}
		
		public void Save( string file_path ) {
			try {
				Console.WriteLine("fastxcel: saving "+file_path);
			
				
				// save xml_shared_Strings...
				SaveSharedStrings();
				
				// save
				foreach ( Worksheet wrk in Worksheets ) {
					wrk.Save();
				}

				//Utils.FileUtils.ReplaceInFile( temp_path + "\\xl\\workbook.xml","\"id\"", "\"r:id\"" );
				Compress( file_path );
			} catch ( Exception ex ) {
				throw new FastXcelException( ex.Message, ex );
			}
		}
		
		private void SaveSharedStrings() {
			string inner_xml_header = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"+
				"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"100000\" uniqueCount=\"99000\">";
			string inner_xml_footer = "</sst>";
			string strgs = String.Empty;
			foreach ( KeyValuePair<int,string> sh_str in shared_strings ) {
				strgs += "<si><t>"+sh_str.Value+"</t></si>";
			}
			if ( xml_shared_strings == null )
				xml_shared_strings = new AmusingXml.XmlFun();
			xml_shared_strings.Document.InnerXml = inner_xml_header + strgs + inner_xml_footer;
			xml_shared_strings.SaveDocument(temp_path+"\\xl\\sharedStrings.xml");
		}
		
		private void ProcessDocumentRelationships() {
			if ( File.Exists(temp_path+"\\xl\\_rels\\workbook.xml.rels") ) {
				xml_wb_rels.LoadDocument(temp_path+"\\xl\\_rels\\workbook.xml.rels");
			} else {
				throw new FastXcelException("Can't find file \\xl\\_rels\\workbook.xml.rels. Document seems to be corrupted.");
			}
		}
		
		private void ProcessWorkbook() {
			if ( File.Exists(temp_path+"\\xl\\workbook.xml") ) {
				xml_workbook.LoadDocument(temp_path+"\\xl\\workbook.xml");
			} else {
				throw new FastXcelException("Can't find file \\xl\\workbook.xml. Document seems to be corrupted.");
			}
		}
		
		private void ReadSharedStrings() {
			if ( File.Exists(temp_path+"\\xl\\sharedStrings.xml") ) {
				xml_shared_strings.LoadDocument(temp_path+"\\xl\\sharedStrings.xml");
				foreach ( XmlElement el in xml_shared_strings.XmlElements ) {
					if ( el.Name == "sst" ) {
						foreach ( XmlNode si_node in el.ChildNodes ) {
							if ( si_node.Name == "si" ) {
								foreach ( XmlNode value_node in si_node.ChildNodes ) {
									if ( value_node.Name == "t" ) {
										shared_strings.Add( shared_strings.Count, value_node.InnerText );
									}
								}
							}
						}
					}
				}
			} else {
				throw new FastXcelException("Can't find file \\xl\\sharedStrings.xml. Document seems to be corrupted.");
			}
		}
		
		private string GetResourceId(string file_path) {
			string ret = String.Empty;
			List<string> tmp_path = new List<string>();
			xml_wb_rels.FillXmlElements(xml_wb_rels.Document);
			foreach ( XmlElement el in xml_wb_rels.FindNodes("Relationship") ) {
				tmp_path = CommonFacilities.Common.ExplodeString(file_path, "\\");
				if ( (el.Attributes["Target"].Value).IndexOf(tmp_path[tmp_path.Count-1]) > -1 ) {
					// Seeking filename in realtionships
					return el.Attributes["Id"].Value;
				}
			}
			return ret;
		}
		
		private string GetWorksheetName(string file_path, string res_id) {
			string ret = String.Empty;
			List<string> tmp_path = new List<string>();
			xml_workbook.FillXmlElements(xml_workbook.Document);
			foreach ( XmlElement el in xml_workbook.FindNodes("sheet") ) {
				tmp_path = CommonFacilities.Common.ExplodeString(file_path, "\\");
				if ( el.Attributes["r:id"].Value == res_id ) {
					// Seeking filename in realtionships
					return el.Attributes["name"].Value;
				}
			}
			return ret;
		}
		
		private short GetWorksheetId(string file_path, string res_id) {
			short ret = -1;
			List<string> tmp_path = new List<string>();
			xml_workbook.FillXmlElements(xml_workbook.Document);
			foreach ( XmlElement el in xml_workbook.FindNodes("sheet") ) {
				tmp_path = CommonFacilities.Common.ExplodeString(file_path, "\\");
				if ( el.Attributes["r:id"].Value.ToString() == res_id ) {
					// Seeking filename in realtionships
					return short.Parse(el.Attributes["sheetId"].Value);
				}
			}
			return ret;
		}
		
		private void ProcessWorksheets() {
			try {
				Worksheets.Clear();
				string[] xl_sheet_files = Directory.GetFiles(temp_path+"\\xl\\worksheets\\", "sheet*.xml");
				
				foreach ( string sheet_file in xl_sheet_files ) {
					string res_id = GetResourceId(sheet_file);
					string name = GetWorksheetName(sheet_file, res_id);
					short sheet_id = GetWorksheetId(sheet_file, res_id);
					Console.WriteLine("---------------------------"+ sheet_file +"-------------------------------------");
					Worksheet wrk =  new Worksheet( shared_strings );
					wrk.Load( res_id, sheet_id, name, sheet_file, shared_strings );
					wrk.Number = (uint)Worksheets.Count;
					Console.WriteLine("---------------------------------------------------------------------------------------");
					Worksheets.Add( wrk );
				}
			} catch ( Exception ex ) {
				throw new FastXcelException( "Error processing worksheets", ex );
			}
		}
		
		public short NewWorksheet( string name = "" ) {
			try {
				
				System.Resources.ResourceManager resourceManager =
					new System.Resources.ResourceManager("fastxcel.Resource", GetType().Assembly);
				
				
				var assembly = System.Reflection.Assembly.GetExecutingAssembly();
				
				short new_ws_number = (short)(Worksheets.Count+1);
				string res_id = "Re";
				
				Worksheet new_wrk = new Worksheet( shared_strings );
				
				if ( name == String.Empty )
					name = "Worksheet"+Worksheets.Count.ToString();
				
				// First of all - create new xml file
				int num = 1;
				while ( File.Exists(temp_path+"\\xl\\worksheets\\sheet"+num.ToString()+".xml") )
					++num;
				string new_file_path = temp_path+"\\xl\\worksheets\\sheet"+num.ToString()+".xml";
				string resource_path = assembly.GetName().Name+".Resources.sheet_xml";
				
				BinaryReader br = new BinaryReader(assembly.GetManifestResourceStream(resource_path));
				long file_length = assembly.GetManifestResourceStream(resource_path).Length;
				byte[] inners = new byte[file_length];
				inners = br.ReadBytes((int)file_length);
				
				br.Close();

				FileStream fs = File.Create( new_file_path );
				BinaryWriter bw = new BinaryWriter( fs );
				
				bw.Write(inners);
				bw.Close();

				
				// Register relationship
				int res_id_num = xml_wb_rels.XmlElements.Count;
				
				AddNewRelationship(num, "rId" + res_id_num.ToString(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
				
				// Add into workbook.xml
				AddSheetInWorkbook( "rId" + res_id_num.ToString(), name, new_ws_number );
				
				// Finally
				new_wrk.Load( new_file_path, shared_strings );

				Worksheets.Add( new_wrk );
				
				return new_ws_number;
			} catch ( Exception ex ) {
				throw new FastXcelException( "Can't create new worksheet"+ex.Message, ex );
			}
			
		}
		
		void AddSheetInWorkbook( string res_id, string ws_name, short sheet_number ) {
			Dictionary<string, string> attrs = new Dictionary<string, string>();
			attrs["name"] = ws_name;
			attrs["r:id"] = res_id;
			attrs["sheetId"] = sheet_number.ToString();
			xml_workbook.FillXmlElements(xml_workbook.Document);
			XmlElement sheets = xml_workbook.FindFirstNode("sheets", new Dictionary<string, string>());

			var new_el = xml_workbook.CreateNode("sheet", attrs, String.Empty, sheets);
			
			
			xml_workbook.Document.InnerXml = xml_workbook.Document.InnerXml.Replace("xmlns=\"\"", String.Empty).Replace("xmlns:r=\"special_ns\"", String.Empty);
			xml_workbook.SaveDocument(temp_path + "\\xl\\workbook.xml");
		}
		
		void AddNewRelationship(int res_number, string res_id, string res_type)
		{
			Dictionary<string, string> attrs = new Dictionary<string, string>();
			attrs["Id"] = res_id;
			attrs["Type"] = res_type;
			attrs["Target"] = "worksheets/sheet" + res_number.ToString() + ".xml";
			XmlElement relships = xml_wb_rels.FindFirstNode("Relationships", new Dictionary<string, string>());
			xml_wb_rels.CreateElement("Relationship", attrs, String.Empty, relships);
			xml_wb_rels.Document.InnerXml = xml_wb_rels.Document.InnerXml.Replace("xmlns=\"\"", String.Empty);
			xml_wb_rels.SaveDocument(temp_path + "\\xl\\_rels\\workbook.xml.rels");
		}
		
		private void CloseDC() {
			try {
				Worksheets.Clear();
				xml_sheets.Clear();
				shared_strings.Clear();
				

				
				xml_workbook = null;
				xml_wb_rels = null;
				xml_shared_strings = null;
				
				GC.Collect();
				GC.WaitForPendingFinalizers();

			} catch ( Exception ex ) {
				throw new FastXcelException( "Error closing a document", ex );
			}
		}
		
		public void Close() {
			try {
				Worksheets.Clear();
				xml_sheets.Clear();
				shared_strings.Clear();
				Directory.Delete( temp_path, true );
			} catch ( Exception ex ) {
				throw new FastXcelException( "Error closing a document", ex );
			}
		}
		
		public void Dispose() {
			Worksheets.Clear();
			xml_sheets.Clear();
			shared_strings.Clear();
		}
	}
}