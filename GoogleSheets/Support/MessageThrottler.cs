using System;
using System.Linq;
using System.Timers;
using System.Collections.Generic;
using System.Reflection;

namespace SheetsPersist
{
	public class MessageThrottler<T> where T: class
	{
		List<string> tabNamesSeenSoFar = new List<string>();
		Timer timer = new Timer();

		object messageLock = new object();
		Dictionary<string, Queue<T>> messages = new Dictionary<string, Queue<T>>();

		DateTime lastBurstTime = DateTime.MinValue;
		readonly TimeSpan minTimeBetweenBursts;
		string defaultSheetName = "No Name";
		string documentName;

		public MessageThrottler(TimeSpan minTimeBetweenBursts)
		{
			this.minTimeBetweenBursts = minTimeBetweenBursts;
			timer.Elapsed += Timer_Elapsed;
			SheetAttribute tabNameAttribute = typeof(T).GetCustomAttribute<SheetAttribute>();
			if (tabNameAttribute != null)
				defaultSheetName = tabNameAttribute.SheetName;
			else
				defaultSheetName = typeof(T).Name;

			DocumentAttribute sheetNameAttribute = typeof(T).GetCustomAttribute<DocumentAttribute>();
			documentName = sheetNameAttribute.DocumentName;
		}

		private void Timer_Elapsed(object sender, ElapsedEventArgs e)
		{
			timer.Enabled = false;
			SendAllMessages();
		}

		void SendAllMessages()
		{
			lastBurstTime = DateTime.Now;
			lock (messageLock)
			{
				foreach (string sheetName in messages.Keys)
				{
					if (messages[sheetName].Any())
					{
						bool firstMessageAdded = false;
						string tabKey = $"{documentName}.{sheetName}";
						if (!tabNamesSeenSoFar.Contains(tabKey))
						{
							firstMessageAdded = GoogleSheets.MakeSureSheetExists<T>(sheetName);
							tabNamesSeenSoFar.Add(tabKey);
						}

						GoogleSheets.InternalAppendRows<T>(messages[sheetName].ToArray(), sheetName, null, firstMessageAdded);
						messages[sheetName].Clear();
					}
				}
			}
		}

		public void AppendRow(T t, string tabName = null)
		{
			if (tabName == null)
				tabName = defaultSheetName;
			
			lock (messageLock)
			{
				if (!messages.ContainsKey(tabName))
					messages[tabName] = new Queue<T>();
				messages[tabName].Enqueue(t);
			}

			// ![](BAEDF4D24FB1C180CE95B77D1FF1A93C.png)

			DateTime now = DateTime.Now;
			TimeSpan timeSinceLastBurst = now - lastBurstTime;
			bool firstBurst = lastBurstTime == DateTime.MinValue;
			if (firstBurst || timeSinceLastBurst > minTimeBetweenBursts)
				SendAllMessages();
			else if (!timer.Enabled)
			{
				DateTime nextScheduledBurstTime = lastBurstTime + minTimeBetweenBursts;
				TimeSpan timeTillNextBurst = nextScheduledBurstTime - now;
				timer.Interval = timeTillNextBurst.TotalMilliseconds;
				timer.Enabled = true;
			}
		}

		public void FlushAllMessages()
		{
			SendAllMessages();
		}
	}
}

