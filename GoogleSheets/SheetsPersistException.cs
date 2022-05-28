using System;
using System.Linq;
using System.Runtime.Serialization;

namespace SheetsPersist
{
	[Serializable]
	public class SheetsPersistException : Exception
	{
		public string DocumentName { get; set; }
		public string SheetName { get; set; }
		// constructors...
		#region SheetsPersistException()
		/// <summary>
		/// Constructs a new SheetsPersistException.
		/// </summary>
		public SheetsPersistException() { }
		#endregion
		#region SheetsPersistException(string message)

		/// <summary>
		/// Constructs a new SheetsPersistException.
		/// </summary>
		/// <param name="message">The exception message</param>
		public SheetsPersistException(string message) : base(message) { }
		#endregion
		#region SheetsPersistException(string message, Exception innerException)
		/// <summary>
		/// Constructs a new SheetsPersistException.
		/// </summary>
		/// <param name="message">The exception message</param>
		/// <param name="innerException">The inner exception</param>
		public SheetsPersistException(string message, Exception innerException) : base(message, innerException) { }
		#endregion
		#region SheetsPersistException(SerializationInfo info, StreamingContext context)
		/// <summary>
		/// Serialization constructor.
		/// </summary>
		protected SheetsPersistException(SerializationInfo info, StreamingContext context) : base(info, context) { }
		#endregion
	}
}

