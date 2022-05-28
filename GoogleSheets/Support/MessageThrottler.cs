using System;
using System.Linq;
using System.Timers;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;

namespace SheetsPersist
{
	public class MessageThrottler<T> where T: class
	{
		List<string> sheetNamesSeenSoFar = new List<string>();
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
			SheetAttribute sheetNameAttribute = typeof(T).GetCustomAttribute<SheetAttribute>();
			if (sheetNameAttribute != null)
				defaultSheetName = sheetNameAttribute.SheetName;
			else
				defaultSheetName = typeof(T).Name;

			DocumentAttribute documentNameAttribute = typeof(T).GetCustomAttribute<DocumentAttribute>();
			documentName = documentNameAttribute.DocumentName;
		}

		private void Timer_Elapsed(object sender, ElapsedEventArgs e)
		{
			timer.Enabled = false;
            SendAllMessages();
		}

		void SendAllMessages()
		{
			lastBurstTime = DateTime.UtcNow;
			lock (messageLock)
			{
				foreach (string sheetName in messages.Keys)
				{
					if (messages[sheetName].Any())
					{
						bool firstMessageAdded = false;
						string tabKey = $"{documentName}.{sheetName}";
						if (!sheetNamesSeenSoFar.Contains(tabKey))
						{
							firstMessageAdded = GoogleSheets.MakeSureSheetExists<T>(sheetName);
							sheetNamesSeenSoFar.Add(tabKey);
						}

						GoogleSheets.InternalAppendRows<T>(messages[sheetName].ToArray(), sheetName, null, firstMessageAdded);
						messages[sheetName].Clear();
					}
				}
			}
		}

        public void AppendRow(T t, string sheetName = null)
		{
			if (sheetName == null)
				sheetName = defaultSheetName;
			
			lock (messageLock)
			{
				if (!messages.ContainsKey(sheetName))
					messages[sheetName] = new Queue<T>();
				messages[sheetName].Enqueue(t);
			}

			// ![](BAEDF4D24FB1C180CE95B77D1FF1A93C.png)

			DateTime now = DateTime.UtcNow;
			TimeSpan timeSinceLastBurst = now - lastBurstTime;
			bool firstBurst = lastBurstTime == DateTime.MinValue;

			if (firstBurst || timeSinceLastBurst > minTimeBetweenBursts)
				SendAllMessages();
			else if (!timer.Enabled)
                EnableTimer(now);
        }

		private void EnableTimer(DateTime now)
        {
            DateTime nextScheduledBurstTime = lastBurstTime + minTimeBetweenBursts;
            TimeSpan timeTillNextBurst = nextScheduledBurstTime - now;
            timer.Interval = timeTillNextBurst.TotalMilliseconds;
            timer.Enabled = true;
        }

        public void FlushAllMessages()
		{
			SendAllMessages();
		}
	}
}

