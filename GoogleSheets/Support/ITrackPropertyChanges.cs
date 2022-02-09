using System;
using System.Linq;
using System.Collections.Generic;

namespace SheetsPersist
{
	public interface ITrackPropertyChanges
	{
		List<string> ChangedProperties { get; set; }
		bool IsDirty { get; set; }
	}
}
