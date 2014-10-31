/*
 * User: Igor.Evdokimov
 * Date: 30.10.2014
 * Time: 18:43
 */
using System;
using System.IO;

namespace fastxcel.Utils
{
	/// <summary>
	/// Description of FileUtils.
	/// </summary>
	public static class FileUtils
	{
		
		public static void ReplaceInFile( string file_path, string source, string target ) {
			FileStream fs = File.Open( file_path, FileMode.Open );
			StreamReader sr = new StreamReader( file_path );
			string content = sr.ReadToEnd();
			sr.Close();
			content.Replace( source, target );
			
			fs = File.Open( file_path, FileMode.Open );
			StreamWriter sw = new StreamWriter(file_path);
			sw.Close();			
		}
	}
}
