(*
Exchange → Apple Calendar synchronization (Delta + 90/30 days)
Only new or changed events since the last run.
Saves sync date + log file to ~/Documents/ExchangeSync/
*)

-- 🗓️ Customize calendar names
set sourceCalendarName to "Exchange"
set targetCalendarName to "Privat"

-- 📅 Period: 30 days back, 30 days forward
set daysBack to 30
set daysAhead to 30

-- 📁 Paths for files in the Documents folder
set docsPath to (POSIX path of (path to documents folder))
set syncDir to docsPath & "ExchangeSync/"
set syncFile to syncDir & "lastSync.txt"
set logFile to syncDir & "sync.log"

-- Make sure the folder exists
do shell script "mkdir -p " & quoted form of syncDir

-- 🕐 Load last sync date or default (30 days back)
set lastSyncDate to missing value
try
	set lastSyncDateString to do shell script "cat " & quoted form of syncFile
	if lastSyncDateString is not "" then
		set lastSyncDate to date lastSyncDateString
	end if
on error
	set lastSyncDate to (current date) - (daysBack * days)
end try

tell application "Calendar"
	try
		set sourceCal to calendar sourceCalendarName
		set targetCal to calendar targetCalendarName
		
		set nowDate to (current date)
		set startDate to nowDate - (daysBack * days)
		set endDate to nowDate + (daysAhead * days)
		
		set sourceEvents to every event of sourceCal whose start date ≥ startDate and start date ≤ endDate
		
		set syncCount to 0
		repeat with e in sourceEvents
			-- Only consider events that were changed or created after the last sync
			set modDate to my safeRead(e, "modification date")
			if modDate is not missing value and modDate > lastSyncDate then
				-- ‼️ Only uncomment when event summaries NEVER contain critical data
				--	set eventSummary to my safeRead(e, "summary")
				set eventStart to my safeRead(e, "start date")
				set eventEnd to my safeRead(e, "end date")
				set eventLocation to my safeRead(e, "location")
				set eventDescription to my safeRead(e, "description")
				set eventRecurrence to my safeRead(e, "recurrence")
				
				if eventStart is not missing value and eventSummary is not "" then
					set existingEvents to (every event of targetCal whose start date = eventStart and summary = eventSummary)
					if (count of existingEvents) = 0 then
						if eventRecurrence is not missing value and (count of eventRecurrence) > 0 then
							make new event at end of events of targetCal with properties {summary:eventSummary, start date:eventStart, end date:eventEnd, location:eventLocation, description:eventDescription, recurrence:eventRecurrence}
						else
							make new event at end of events of targetCal with properties {summary:eventSummary, start date:eventStart, end date:eventEnd, location:eventLocation, description:eventDescription}
						end if
						set syncCount to syncCount + 1
					end if
				end if
			end if
		end repeat
		
		-- 🔄 Update sync date
		do shell script "echo " & quoted form of (nowDate as text) & " > " & quoted form of syncFile
		
		-- 🪵 Write Logs
		set logEntry to (do shell script "date '+%Y-%m-%d %H:%M:%S'") & " | " & syncCount & " new appointments synchronized"
		do shell script "echo " & quoted form of logEntry & " >> " & quoted form of logFile
		
		display notification "Delta sync completed: " & syncCount & " new appointments copied ✅"with title "Exchange → Apple Calendar"
		
	on error errMsg number errNum
		set logEntry to (do shell script "date '+%Y-%m-%d %H:%M:%S'") & " | ERROR (" & errNum & "): " & errMsg
		do shell script "echo " & quoted form of logEntry & " >> " & quoted form of logFile"
		display dialog "Sync error: " & errMsg & " (Error number: " & errNum & ")" buttons {"OK"} default button "OK"
	end try
end tell


-- 🔧 Safe reading
on safeRead(eventObj, propName)
	try
		tell application "Calendar"
			if propName is "summary" then return my safeText(summary of eventObj)
			if propName is "start date" then return start date of eventObj
			if propName is "end date" then return end date of eventObj
			if propName is "location" then return my safeText(location of eventObj)
			if propName is "description" then return my safeText(description of eventObj)
			if propName is "recurrence" then return recurrence of eventObj
			if propName is "modification date" then return modification date of eventObj
		end tell
	on error
		return missing value
	end try
end safeRead

-- 🔧 Safe text
on safeText(theValue)
	if theValue is missing value then
		return ""
	else
		return theValue as text
	end if
end safeText
