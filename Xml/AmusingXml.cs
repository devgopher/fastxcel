/*
 * Сделано в SharpDevelop.
 * Пользователь: Igor.Evdokimov
 * Дата: 17.09.2012
 * Время: 17:33
 * Класс для упрощенной работы с XML
 * Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
 */
using System;
using System.Xml;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AmusingXml
{
	public class ExtXmlNode {
		
		public XmlNode InnerNode { get; private set; }
		
		
		public ExtXmlNode( XmlNode base_el ) {
			InnerNode = base_el;
			ParentNode = base_el.ParentNode;
			Tag = String.Empty;
			parallel_options.MaxDegreeOfParallelism = 4;
		}

		public String Tag {get; set;}

		public XmlNode ParentNode {get; set;}
		
		ParallelOptions parallel_options= new ParallelOptions();

		
		/// <summary>
		/// Поиск подузлов по заданному аттрибуту
		/// </summary>
		/// <param name="_conds"></param>
		/// <returns></returns>
		public List<ExtXmlNode> FindSubElementsByAttribute( KeyValuePair<string, string> _conds ) {
			List<ExtXmlNode> ret = new List<ExtXmlNode>();
			Parallel.ForEach( (this.InnerNode.ChildNodes).Cast<XmlElement>(), parallel_options, (sub_el, loop_state) => {
			                 	if ( sub_el.HasAttribute(_conds.Key)  == true ) {
			                 		if ( sub_el.Attributes[_conds.Key].Value == _conds.Value ) {
			                 			lock (ret) {
			                 				ret.Add(new ExtXmlNode(sub_el));
			                 			}
			                 		}
			                 	}
			                 } );
			return ret;
		}
		
		
		/// <summary>
		/// Поиск первого подузла по заданному аттрибуту
		/// </summary>
		/// <param name="_conds"></param>
		/// <returns></returns>
		public ExtXmlNode FindFirstSubElementByAttribute(string _name, string _value) {
			ExtXmlNode ret = null;
			Parallel.ForEach( (this.InnerNode.ChildNodes).Cast<ExtXmlNode>(), parallel_options, (sub_el, loop_state) => {
			                 	if ( sub_el.InnerNode is XmlElement )
			                 		if ( ((XmlElement)sub_el.InnerNode).HasAttribute(_name)  == true ) {
			                 		if ( sub_el.InnerNode.Attributes[_name].Value ==_value ) {
			                 			ret = sub_el;
			                 			loop_state.Break();
			                 		}
			                 	}
			                 } );
			return ret;
		}
	}
	
	public class XmlFunException : Exception
	{
		public XmlFunException(string message) : base(message) {}
	}
	
	/// <summary>
	/// Description of AmusingXml.
	/// </summary>
	public class XmlFun
	{
		public XmlFun()
		{
			XmlElements = new List<XmlElement>();
			parallel_options.MaxDegreeOfParallelism = 4;
		}
		
		ParallelOptions parallel_options= new ParallelOptions();

		
		/// <summary>
		/// Создаем документ и назначаем ему имя
		/// </summary>
		/// <param name="document_name">имя</param>
		public void CreateDocument( string document_name ) {
			xml_document = new XmlDocument();
		}
		
		
		/// <summary>
		/// Создаем документ и назначаем ему имя
		/// </summary>
		/// <param name="document_name">имя</param>
		/// <param name="inner_xml">содержимое</param>
		public void CreateDocument( string document_name, string inner_xml ) {
			CreateDocument( document_name );
			xml_document.InnerXml = inner_xml;
			XmlElements.Clear();
			FillXmlElements( xml_document );
		}
		
		public void SaveDocument( string new_path = null ) {
			xml_document.Save( new_path );
		}
		
		
		public int _SortFunc( XmlElement x1, XmlElement x2 ) {
			try {
				if ( x1 != null && x2 != null ) {
					if ( x1.Name != null && x2.Name != null )
						if (x1.Name.CompareTo(x2.Name) >= 2)
							return 2;
				}
			} catch ( Exception ) {
				// Doing nothing
			}
			return 1;
		}
		
		public void FillXmlElements( XmlDocument doc ) {
			XmlElements.Clear();
			XmlElements.Add(doc.DocumentElement);
			Parallel.ForEach( doc.DocumentElement.ChildNodes.Cast<XmlNode>(), parallel_options, (el, loop_state) => {
			                 	if ( ! (el is XmlDeclaration) && !(el is XmlSignificantWhitespace) ) {
			                 		ExtXmlNode ext_el =  new ExtXmlNode(el);
			                 		if ( ext_el != null ) {
			                 			lock ( XmlElements ) {
			                 				XmlElements.Add((XmlElement)ext_el.InnerNode);
			                 			}
			                 			FillXmlElements(ext_el);
			                 		}
			                 	}
			                 });
			
			XmlElements.Sort(_SortFunc);
		}
		
		public void FillXmlElements( ExtXmlNode elem ) {
			Parallel.ForEach( elem.InnerNode.ChildNodes.Cast<XmlNode>(), parallel_options, (el, loop_state) => {
			                 	ExtXmlNode ext_el =  new ExtXmlNode(el);
			                 	if ( ext_el != null ) {
			                 		if ( !(ext_el.InnerNode is XmlText)
			                 		    &&  !(ext_el.InnerNode is XmlDeclaration)
			                 		    &&  !(ext_el.InnerNode is XmlWhitespace)
			                 		    &&  !(ext_el.InnerNode is XmlSignificantWhitespace) ) {
			                 			lock (XmlElements) {
			                 				XmlElements.Add((XmlElement)ext_el.InnerNode);
			                 			}
			                 			FillXmlElements(ext_el);
			                 		}
			                 	}
			                 } );
		}
		
		/// <summary>
		/// сортируем узлы по значению аттрибута
		/// </summary>
		/// <param name="node_name"></param>
		/// <param name="attr_name"></param>
		public void SortNodesByAttrValue( string node_name, string attr_name ) {
			var row_nodes = ExtFindNodes( node_name);
			if ( row_nodes.Count > 0 ) {
				row_nodes.Sort( (ExtXmlNode el1, ExtXmlNode el2) => {
				               	return (el1.InnerNode.Attributes[attr_name].Value.CompareTo(el2.InnerNode.Attributes[attr_name].Value));
				               } );
				// удаляем узлы, упорядоченные по-старому...
				foreach (var nd in row_nodes) {
					nd.ParentNode.RemoveChild(nd.InnerNode);
				}
				
				// вставляем в новом порядке
				XmlNode prev_node = null;
				foreach (var nd in row_nodes) {
					nd.ParentNode.InsertBefore(nd.InnerNode, null);
					prev_node = nd.InnerNode;
				}
			}
		}
		
		/// <summary>
		/// Делегат для сортировки узлов
		/// </summary>
		public delegate int NodeSortFunc(ExtXmlNode el1, ExtXmlNode el2, string attr_name);
		
		/// <summary>
		/// сортируем узлы по значению аттрибута
		/// </summary>
		/// <param name="node_name"></param>
		/// <param name="attr_name"></param>
		public void SortNodesByAttrValue( string node_name, string attr_name, NodeSortFunc _sort_delegate ) {
			var row_nodes = ExtFindNodes( node_name );

			if ( row_nodes.Count > 0 ) {
				row_nodes.Sort( (ExtXmlNode el1, ExtXmlNode el2) => { return _sort_delegate(el1,el2,attr_name);} );
				// удаляем узлы, упорядоченные по-старому...
				foreach (var nd in row_nodes) {
					if ( nd.ParentNode != null )
						nd.ParentNode.RemoveChild(nd.InnerNode);
				}
				
				// вставляем в новом порядке
				XmlNode prev_node = null;
				foreach (var nd in row_nodes) {
					if ( nd.ParentNode != null ) {
						nd.ParentNode.InsertBefore(nd.InnerNode, null);
						prev_node = nd.InnerNode;
					}
				}
			}
		}
		
		/// <summary>
		/// сортируем узлы по значению аттрибута, полагая, что значение типа integer
		/// </summary>
		/// <param name="node_name"></param>
		/// <param name="attr_name"></param>
		public void SortNodesByAttrIntValue( string node_name, string attr_name ) {
			var row_nodes = FindNodes( node_name, new Dictionary<string, string>() );
			XmlNode global_parent = FindFirstNode( "sheetData", new Dictionary<string, string>());
			XmlNode ws_node = FindFirstNode( "worksheet", new Dictionary<string, string>());
			XmlNode sh_format_pr = FindFirstNode( "sheetFormatPr", new Dictionary<string, string>());
			
			if ( row_nodes.Count > 0 ) {
				row_nodes.Sort( (XmlElement el1, XmlElement el2) => {
				               	try {
				               		if (int.Parse(el1.Attributes[attr_name].Value) == int.Parse(el2.Attributes[attr_name].Value))
				               			return 0;
				               		else if (int.Parse(el1.Attributes[attr_name].Value) > int.Parse(el2.Attributes[attr_name].Value))
				               			return 1;
				               		else
				               			return -1;
				               	} catch ( Exception ex ) {
				               		return int.MinValue;
				               	}
				               }	);
				// удаляем узлы, упорядоченные по-старому...
				foreach ( XmlNode subn in global_parent.ChildNodes )
					global_parent.RemoveChild(subn);
				
				ws_node.RemoveChild(global_parent);
				
				
				// вставляем в новом порядке
				XmlNode prev_node = null;
				foreach (var nd in row_nodes) {
					//if ( nd.ParentNode != null )
					global_parent.InsertAfter(nd, prev_node);
					prev_node = nd;
				}
				
				ws_node.InsertAfter( global_parent, sh_format_pr );
			}
		}
		
		/// <summary>
		/// Создаем узел с аттрибутами
		/// </summary>
		/// <param name="name"></param>
		/// <param name="attributes"></param>
		/// <param name="inner_text"></param>
		/// <param name="parent_node"></param>
		public XmlElement CreateElement( string name, Dictionary<string, string> attributes,
		                                string inner_text, XmlElement parent_element ) {
			try{
				XmlElement node = xml_document.CreateElement(name);
				// пройдемся по аттрибутам...
				if (attributes != null) {
					foreach ( KeyValuePair<string, string>param in attributes ) {
						XmlAttribute attrib = null;
						if ( param.Key.IndexOf(":")>0 )
						{
							attrib = xml_document.CreateAttribute(param.Key.Substring(0, param.Key.IndexOf(":")),
							                                      param.Key.Substring(param.Key.IndexOf(":")+1, param.Key.Length-param.Key.IndexOf(":")-1));
						} else
							attrib = xml_document.CreateAttribute(param.Key);
						attrib.InnerText = param.Value.ToString();
						node.Attributes.Append(attrib);
					}
				}
				node.InnerText = inner_text;
				parent_element.AppendChild(node);
				XmlElements.Add(node);
				return node;
			} catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
		}
		
		/// <summary>
		/// Создаем узел с аттрибутами
		/// </summary>
		/// <param name="name"></param>
		/// <param name="attributes"></param>
		/// <param name="inner_text"></param>
		/// <param name="parent_node"></param>
		public XmlNode CreateNode( string name, Dictionary<string, string> attributes,
		                          string inner_text, XmlNode parent_node ) {
			try{
				XmlNode node = xml_document.CreateNode(XmlNodeType.Element, name, String.Empty);
				// пройдемся по аттрибутам...
				if (attributes != null) {
					foreach ( KeyValuePair<string, string>param in attributes ) {
						XmlAttribute attrib = null;
						if ( param.Key.IndexOf(":")>0 )
						{
							attrib = xml_document.CreateAttribute(param.Key, "special_ns");
						} else
							attrib = xml_document.CreateAttribute(param.Key);
						attrib.InnerText = param.Value.ToString();
						node.Attributes.Append(attrib);
					}
				}
				node.InnerText = inner_text;
				parent_node.AppendChild(node);
				
				XmlElements.Add((XmlElement)node);
				return node;
			} catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
		}
		
		/// <summary>
		/// Удаление узла
		/// </summary>
		/// <param name="name">Имя узла</param>
		/// <param name="attributes">Аттрибуты зла</param>
		public void RemoveNode( string name, Dictionary<string, string> attributes = null) {
			XmlElement tmp_el = FindFirstNode( name, attributes );
			tmp_el.RemoveAll();
			if ( tmp_el != null ) {
				if ( tmp_el.ParentNode != null )
					tmp_el.ParentNode.RemoveChild(tmp_el);
				else
					tmp_el.OwnerDocument.RemoveChild(tmp_el);
				FillXmlElements(xml_document);
				XmlElements.Remove(tmp_el);
			}
			
			FillXmlElements(xml_document);
		}
		
		public List<ExtXmlNode> ExtFindNodes( string name, Dictionary<string, string> attributes = null) {
			List<ExtXmlNode> result_nodes = new List<ExtXmlNode>();
			List<XmlElement> tmp_nodes = null;
			if ( attributes != null ) {
				if ( attributes.Count > 0 ) {
					tmp_nodes = FindNodes( attributes );
				}	else {
					tmp_nodes = XmlElements;
				}
			} else
				tmp_nodes = XmlElements;
			
			Parallel.ForEach( tmp_nodes, parallel_options, (el, loop_state) => {
			                 	if ( el.Name == name ) {
			                 		lock (result_nodes) {
			                 			result_nodes.Add(new ExtXmlNode(el));
			                 		}
			                 	}
			                 } );
			return result_nodes;
		}
		
		public List<XmlElement> FindNodes( string name, Dictionary<string, string> attributes = null) {
			List<XmlElement> result_nodes = new List<XmlElement>();
			List<XmlElement> tmp_nodes = null;
			if ( attributes != null ) {
				if ( attributes.Count > 0 ) {
					tmp_nodes = FindNodes( attributes );
				}	else {
					tmp_nodes = XmlElements;
				}
			} else
				tmp_nodes = XmlElements;
			
			Parallel.ForEach( tmp_nodes, parallel_options, (el, loop_state) => {
			                 	if ( el.Name == name ) {
			                 		lock (result_nodes) {
			                 			result_nodes.Add(el);
			                 		}
			                 	}
			                 } );
			return result_nodes;
		}
		
		/// <summary>
		/// Ищем узлы по аттрибутам
		/// </summary>
		/// <param name="attributes"></param>
		/// <returns></returns>
		public List<XmlElement> FindNodes( Dictionary<string, string> attributes ) {
			List<XmlElement> result_nodes = new List<XmlElement>();
			try {

				if ( attributes != null ) {
					Parallel.ForEach( XmlElements, parallel_options, (el, loop_state) => {
					                 	bool decision = true; // Проверяем все параметры и если они все совпадают, то true

					                 	SearchNodes( el, ref decision, attributes, false);
					                 	// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
					                 	if (decision == true) {
					                 		lock (result_nodes) {
					                 			result_nodes.Add(el);
					                 		}
					                 	}
					                 } );
				} else {
					result_nodes = XmlElements;
				}
			}  catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
			
			return result_nodes;
		}

		void SearchNodes(XmlElement el, ref bool decision, Dictionary<string, string> attributes, bool only_first_match = false)
		{
			foreach (KeyValuePair<string, string> attrs in attributes) {
				try {
					if (el.Attributes[attrs.Key] == null) {
						// Нет такого аттрибута
						decision = false;
					} else {
						// аттрибут есть, но значение не совпадает
						if (el.Attributes[attrs.Key].Value != attrs.Value) {
							decision = false;
							// Если значение аттрибута не совпало, то смысл дальше искать?
							break;
						}
					}
				} catch ( Exception )  {
					decision = false;
				}
				
				// Если ищем только первое вхождение и нашли его - выход!
				if (decision == true && only_first_match == true)
					return;
			}
		}
		
		/// <summary>
		/// Ищем узел (первый попавшийся) по аттрибутам и имени
		/// </summary>
		/// <param name="attributes"></param>
		/// <returns></returns>
		public XmlElement FindFirstNode(string name, Dictionary<string, string> attributes) {
			try {
				XmlElement ret = null;
				Parallel.ForEach( XmlElements, parallel_options, (el, loop_state) => {
				                 	if ( el.Name == name ) {
				                 		if ( attributes != null ) {
				                 			if ( attributes.Count > 0 ) {
				                 				bool decision = true; // Проверяем все параметры и если они все совпадают, то true
				                 				SearchNodes(el, ref decision, attributes, true);
				                 				// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
				                 				if (decision == true) {
				                 					ret = el;
				                 					loop_state.Break();
				                 				} else
				                 					ret = el;
				                 			} else
				                 				ret = el;
				                 		}
				                 	}
				                 } );
				return ret;
			}  catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
		}
		
		/// <summary>
		/// Ищем узел (первый попавшийся) по аттрибутам
		/// </summary>
		/// <param name="attributes"></param>
		/// <returns></returns>
		public XmlElement FindFirstNode(Dictionary<string, string> attributes) {
			try {
				XmlElement ret = null;
				Parallel.ForEach( XmlElements, parallel_options, (el, loop_state) => {
				                 	bool decision = true; // Проверяем все параметры и если они все совпадают, то true
				                 	SearchNodes(el, ref decision, attributes, true);
				                 	// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
				                 	if (decision == true) {
				                 		ret = el;
				                 		loop_state.Break();
				                 	}
				                 } );
				return ret;
			}  catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
		}
		
		/// <summary>
		/// Загрузка документа
		/// </summary>
		/// <param name="filename"></param>
		public void LoadDocument( string filename ) {
			try {
				XmlTextReader xmlreader = new XmlTextReader( File.Open( filename, FileMode.Open ));
				xml_document = new XmlDocument();
				xml_document.Load( xmlreader );
				XmlElements.Clear();
				FillXmlElements( xml_document );
				xmlreader.Close();
			} catch ( XmlException ex ) {
				throw new XmlFunException(ex.Message);
			}
		}
		
		/// <summary>
		/// Загрузка документа
		/// </summary>
		/// <param name="filename"></param>
		public void LoadDocument( XmlDocument _xml_document ) {
			xml_document = _xml_document;
			XmlElements.Clear();
			FillXmlElements( xml_document );
		}
		
		/// <summary>
		/// Получаем результирующий текст документа
		/// </summary>
		/// <returns></returns>
		public string GetText() {
			return xml_document.InnerXml;
		}
		
		/// <summary>
		/// Текст в формате "ключ1=значение1,ключ2=значение2"
		/// </summary>
		/// <param name="text">текст</param>
		/// <returns></returns>
		public Dictionary<string, string> TextToDictionary(string text) {
			try {
				if (text != null) {
					List<string> strgs = CommonFacilities.Common.ExplodeString(text.Trim(), ",");
					Dictionary<string, string> dict = new Dictionary<string, string>();
					foreach (string strg in strgs) {
						List<string> tmp_key_val = CommonFacilities.Common.ExplodeString(strg,"=");
						dict.Add(tmp_key_val[0], tmp_key_val[1]);
					}
					return dict;
				}
				return null;
			} catch (Exception ex) {
				throw new XmlFunException(ex.Message);
			}
		}

		
		public XmlDocument Document{get {return xml_document;} private set {xml_document = value;}}
		
		public List<XmlElement> XmlElements{get;set;}
		
		private XmlDocument xml_document;
		
	}
}