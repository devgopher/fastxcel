/*
 * Пользователь: Igor.Evdokimov
 * Дата: 31.07.2014
 * Время: 12:11
 */
using System;
using System.Runtime.Serialization;

namespace fastxcel
{
	/// <summary>
	/// Desctiption of FastXcelException.
	/// </summary>
	public class FastXcelException : Exception, ISerializable
	{
		public FastXcelException()
		{
		}

	 	public FastXcelException(string message) : base(message)
		{
		}

		public FastXcelException(string message, Exception innerException) : base(message, innerException)
		{
		}

		// This constructor is needed for serialization.
		protected FastXcelException(SerializationInfo info, StreamingContext context) : base(info, context)
		{
		}
	}
}