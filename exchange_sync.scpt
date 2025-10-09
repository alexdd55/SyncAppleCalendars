(*
Exchange → Apple Calendar Synchronization (Delta + 90/30 days)
Copies only new or modified events since the last run.
Stores sync timestamp + log file in ~/Documents/ExchangeSync/
*)

-- 🗓️ Calendar names (adjust as needed)
set sourceCalendarName to "Exchange"
set targetCalendarName to "MyCalender"

-- 📅 Time range: X days back, Y days ahead
set daysBack to 5
set daysAhead to 35

-- 📁 File paths in user's Documents folder
set docsPath to (POSIX path of (path to documents folder))
set syncDir to docsPath & "ExchangeSync/"
set syncFile to syncDir & "lastSync.txt"
set logFile to syncDir & "sync.log"

-- Ensure sync directory exists
do shell script "mkdir -p " & quoted form of syncDir

-- 🕐 Load last sync date (or use default = 90 days ago)
set lastSyncDate to (current date) - (daysBack * days)
try
	set lastSyncDateString to do shell script "cat " & quoted form of syncFile
	if lastSyncDateString is not "" then set lastSyncDate to my parseISODate(lastSyncDateString)
on error
	set lastSyncDate to (current date) - (daysBack * days)
end try

-- 🕑 Define sync window
set nowDate to current date
set startDate to nowDate - (daysBack * days)
set endDate to nowDate + (daysAhead * days)

-- Counter for new events
set syncCount to 0

try
	-- All Calendar-related commands must be inside this block
	tell application "Calendar"
		set sourceCal to calendar sourceCalendarName
		set targetCal to calendar targetCalendarName
		
		-- Fetch all events within the time range
		set sourceEvents to every event of sourceCal whose start date ≥ startDate and start date ≤ endDate
		
		repeat with e in sourceEvents
			-- Safely read modification date (handles Exchange exceptions)
			set modDate to my safeRead(e, "modification date")
			
			-- Only process events modified after last sync
			if (modDate is not missing value) and (class of modDate is date) and (modDate > lastSyncDate) then
				-- Read and sanitize event properties
				set eventSummary to my safeText(my safeRead(e, "summary"))
				set eventStart to my safeRead(e, "start date")
				set eventEnd to my safeRead(e, "end date")
				set eventLocation to my safeText(my safeRead(e, "location"))
				set eventRecurrence to my safeRead(e, "recurrence")
				
				-- Ensure event has required data
				if eventStart is not missing value and eventSummary is not "" then
					set existingEvents to (every event of targetCal whose start date = eventStart and end date = eventEnd and summary = eventSummary)
					
					-- Create event only if it doesn't exist yet
					if (count of existingEvents) = 0 then
						try
							if eventRecurrence is not missing value and (count of eventRecurrence) > 0 then
								make new event at end of events of targetCal with properties {summary:eventSummary, start date:eventStart, end date:eventEnd, location:eventLocation, recurrence:eventRecurrence}
							else
								make new event at end of events of targetCal with properties {summary:eventSummary, start date:eventStart, end date:eventEnd, location:eventLocation}
							end if
							set syncCount to syncCount + 1
						on error innerErrMsg number innerErrNum
							my appendToFile(my buildLogEntry("Inner error (" & innerErrNum & "): " & innerErrMsg), logFile)
						end try
					end if
				end if
			end if
		end repeat
	end tell
	
	-- 🔄 Save current sync time in ISO format
	set isoDate to do shell script "date -u '+%Y-%m-%dT%H:%M:%SZ'"
	do shell script "echo " & quoted form of isoDate & " > " & quoted form of syncFile
	
	-- 🪵 Write sync summary to log
	my appendToFile(my buildLogEntry((syncCount as text) & " new events synchronized"), logFile)
	
	-- ✅ User notification
	my notifySuccess(syncCount)
	
on error errMsg number errNum
	-- Handle outer-level errors (e.g., Calendar not available)
	my appendToFile(my buildLogEntry("ERROR (" & errNum & "): " & errMsg), logFile)
	my notifyError(errMsg, errNum)
end try


------------------------------------------------------------
-- 🧩 Utility functions
------------------------------------------------------------

-- Safe reading of Calendar event properties
on safeRead(eventObj, propName)
	tell application "Calendar"
		try
			if propName is "summary" then
				return summary of eventObj
			else if propName is "start date" then
				return start date of eventObj
			else if propName is "end date" then
				return end date of eventObj
			else if propName is "location" then
				return location of eventObj
			else if propName is "recurrence" then
				try
					return recurrence of eventObj
				on error
					return missing value
				end try
			else if propName is "modification date" then
				try
					set modDate to modification date of eventObj
					if modDate is missing value then error "No modification date"
					return modDate
				on error
					try
						-- Fallback to start date if modification date is unavailable
						set fallbackDate to start date of eventObj
						if fallbackDate is missing value then set fallbackDate to ((current date) - (90 * days))
						return fallbackDate
					on error
						return ((current date) - (90 * days))
					end try
				end try
			else
				return missing value
			end if
		on error
			return missing value
		end try
	end tell
end safeRead

-- Convert ISO-8601 string → AppleScript date
on parseISODate(isoString)
	try
		set y to text 1 thru 4 of isoString
		set m to text 6 thru 7 of isoString
		set d to text 9 thru 10 of isoString
		set hh to text 12 thru 13 of isoString
		set mm to text 15 thru 16 of isoString
		set ss to text 18 thru 19 of isoString
		
		set theDate to (current date)
		set year of theDate to y as integer
		set month of theDate to m as integer
		set day of theDate to d as integer
		set time of theDate to ((hh as integer) * hours) + ((mm as integer) * minutes) + (ss as integer)
		return theDate
	on error
		return (current date)
	end try
end parseISODate

-- Convert safely any value to text
on safeText(theValue)
	if theValue is missing value then
		return ""
	else
		try
			return theValue as text
		on error
			return (theValue as string)
		end try
	end if
end safeText

-- Append a line to the log file
on appendToFile(theText, thePath)
	do shell script "echo " & quoted form of theText & " >> " & quoted form of thePath
end appendToFile

-- Build timestamped log entry
on buildLogEntry(msg)
	set timestamp to do shell script "date '+%Y-%m-%d %H:%M:%S'"
	return timestamp & " | " & msg
end buildLogEntry

-- macOS user notifications
on notifySuccess(syncCount)
	display notification "Delta sync complete: " & syncCount & " new events copied ✅" with title "Exchange → Apple Calendar"
end notifySuccess

on notifyError(errMsg, errNum)
	display notification "Error (" & errNum & "): " & errMsg with title "Exchange Sync Error"
end notifyError
