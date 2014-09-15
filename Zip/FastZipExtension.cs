/*
 * Пользователь: Igor.Evdokimov
 * Дата: 01.09.2014
 * Время: 12:51
 */
using System;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using System.Collections.Generic;

namespace fastxcel.Zip
{
	/// <summary>
	/// Description of FastZipExtension.
	/// </summary>
	public partial class FastZip : ICSharpCode.SharpZipLib.Zip.FastZip
	{
		public FastZip()
		{
			
		}
		
		/// <summary>
		/// Extract ZIP to memory
		/// </summary>
		/// <param name="inputStream"></param>
		public List<ZipEntry> ExtractZip(Stream inputStream)
		{
			using (ZipFile zip_file = new ZipFile(inputStream)) {
				System.Collections.IEnumerator enumerator = zip_file.GetEnumerator();
				while ( enumerator.MoveNext()) {
					ZipEntry entry = (ZipEntry)enumerator.Current;
					
					if (entry.IsFile)
					{
						ExtractEntry(entry);
					//	ZipOutputStream zos = ((ZipFile)entry).GetOutputStream( entry );
					}
					else if (entry.IsDirectory) {
						ExtractEntry(entry);
					}
				}
			}
		}
	}
}
