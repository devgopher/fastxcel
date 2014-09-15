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
using System.Collections.Generic;

namespace AmusingXml
{
	public class ExtXmlNode {
		
		public XmlNode InnerNode { get; private set; }
		
		
		
		public ExtXmlNode( XmlNode base_el ) {
			InnerNode = base_el;
			ParentNode = base_el.ParentNode;
			Tag = String.Empty;
		}

		
		/*	public ExtXmlElement(string prefix, string localName, string namespaceURI, XmlDocument doc) {
			
			InnerElement = new XmlElement( prefix, localName, namespaceURI, doc );
			
			Tag = String.Empty;
		}
		 */
		public String Tag {get; set;}

		public XmlNode ParentNode {get; set;}
		
		/// <summary>
		/// Поиск подузлов по заданному аттрибуту
		/// </summary>
		/// <param name="_conds"></param>
		/// <returns></returns>
		public List<ExtXmlNode> FindSubElementsByAttribute( KeyValuePair<string, string> _conds ) {
			List<ExtXmlNode> ret = new List<ExtXmlNode>();
			foreach( XmlElement sub_el in this.InnerNode.ChildNodes ) {
				if ( sub_el.HasAttribute(_conds.Key)  == true ) {
					if ( sub_el.Attributes[_conds.Key].Value == _conds.Value ) {
						ret.Add(new ExtXmlNode(sub_el));
					}
				}
			}
			return ret;
		}
		
		
		/// <summary>
		/// Поиск первого подузла по заданному аттрибуту
		/// </summary>
		/// <param name="_conds"></param>
		/// <returns></returns>
		public ExtXmlNode FindFirstSubElementByAttribute(string _name, string  _value) {
			foreach( ExtXmlNode sub_el in this.InnerNode.ChildNodes ) {
				if ( sub_el.InnerNode is XmlElement )
					if ( ((XmlElement)sub_el.InnerNode).HasAttribute(_name)  == true ) {
					if ( sub_el.InnerNode.Attributes[_name].Value ==_value ) {
						return sub_el;
					}
				}
			}
			return null;
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
		}
		
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
		
		public void SaveDocument( string path = null ) {
			xml_document.Save( path );
		}
		
		public void FillXmlElements( XmlDocument doc ) {
			XmlElements.Clear();
			XmlElements.Add(doc.DocumentElement);
			foreach( XmlNode el in doc.DocumentElement.ChildNodes ) {
				if ( ! (el is XmlDeclaration) ) {
					ExtXmlNode ext_el =  new ExtXmlNode(el);
					if ( ext_el != null ) {
						XmlElements.Add((XmlElement)ext_el.InnerNode);
						FillXmlElements(ext_el);
					}
				}
			}
			
			XmlElements.Sort(( XmlElement x1, XmlElement x2 ) => {
			                 	try {
			                 		if ( x1 != null && x2 != null ) {
			                 			if ( x1.Name != null && x2.Name != null )
			                 				if (x1.Name.CompareTo(x2.Name) >= 2)
			                 					return 2;
			                 		}
			                 	} catch ( Exception ex ) {
			                 		// Doing nothing
			                 	}
			                 	return 1;
			                 }
			                );
		}
		
		public void FillXmlElements( ExtXmlNode elem ) {
			foreach( XmlNode el in elem.InnerNode.ChildNodes ) {
				ExtXmlNode ext_el =  new ExtXmlNode(el);
				if ( ext_el != null ) {
					if ( !(ext_el.InnerNode is XmlText) &&  !(ext_el.InnerNode is XmlDeclaration) &&  !(ext_el.InnerNode is XmlWhitespace) ) {
						XmlElements.Add((XmlElement)ext_el.InnerNode);
						FillXmlElements(ext_el);
					}
				}
			}
		}
		
		/// <summary>
		/// сортируем узлы по значению аттрибута
		/// </summary>
		/// <param name="node_name"></param>
		/// <param name="attr_name"></param>
		public void SortNodesByAttrValue( string node_name, string attr_name ) {
			var row_nodes = ExtFindNodes( node_name);
			XmlNode save_par_node = null;
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
				               	if (int.Parse(el1.Attributes[attr_name].Value) > int.Parse(el2.Attributes[attr_name].Value)) {
				               		return 1;
				               	} else
				               		return -1;
				               } );
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
		public XmlElement CreateNode( string name, Dictionary<string, string> attributes,
		                             string inner_text, XmlElement parent_node ) {
			try{
				XmlElement node = xml_document.CreateElement(name);
				// пройдемся по аттрибутам...
				if (attributes != null) {
					foreach ( KeyValuePair<string, string>param in attributes ) {
						XmlAttribute attrib = xml_document.CreateAttribute(param.Key);
						attrib.InnerText = param.Value.ToString();
						node.Attributes.Append(attrib);
					}
				}
				
				parent_node.AppendChild((XmlElement)node);
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
			List<XmlElement> tmp_nodes = FindNodes( attributes );
			
			foreach ( XmlElement el in tmp_nodes ) {
				if ( el.Name == name )
					result_nodes.Add(new ExtXmlNode(el));
			}
			return result_nodes;
		}
		
		public List<XmlElement> FindNodes( string name, Dictionary<string, string> attributes = null) {
			List<XmlElement> result_nodes = new List<XmlElement>();
			List<XmlElement> tmp_nodes = FindNodes( attributes );
			
			foreach ( XmlElement el in tmp_nodes ) {
				if ( el.Name == name )
					result_nodes.Add(el);
			}
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
					foreach (XmlElement el in XmlElements) {
						bool decision = true; // Проверяем все параметры и если они все совпадают, то true

						SearchNodes( el, ref decision, attributes, false);
						// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
						if (decision == true)
							result_nodes.Add(el);
					}
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
				
				foreach (XmlElement el in XmlElements) {
					if ( el.Name == name ) {
						bool decision = true; // Проверяем все параметры и если они все совпадают, то true
						SearchNodes(el, ref decision, attributes, true);
						// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
						if (decision == true)
							return el;
					}
				}
			}  catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
			return null;
		}
		
		/// <summary>
		/// Ищем узел (первый попавшийся) по аттрибутам
		/// </summary>
		/// <param name="attributes"></param>
		/// <returns></returns>
		public XmlElement FindFirstNode(Dictionary<string, string> attributes) {
			try {
				foreach (XmlElement el in XmlElements) {
					bool decision = true; // Проверяем все параметры и если они все совпадают, то true
					SearchNodes(el, ref decision, attributes, true);
					// Если решение положительно (полное совпадение по условиям), то добавляем узел в результат
					if (decision == true)
						return el;
				}
			}  catch (XmlException ex) {
				throw new XmlFunException(ex.Message);
			}
			return null;
		}
		
		/// <summary>
		/// Загрузка документа
		/// </summary>
		/// <param name="filename"></param>
		public void LoadDocument( string filename ) {
			try {
				xml_document = new XmlDocument();
				xml_document.Load(filename);
				XmlElements.Clear();
				FillXmlElements( xml_document );
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

		public XmlDocument Document{get {return xml_document;}}
		
		public List<XmlElement> XmlElements{get;set;}
		
		private XmlDocument xml_document;
	}
}
