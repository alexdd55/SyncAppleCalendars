(*
Exchange → Apple Calendar Sync
UID-based · Delta · Update · Delete
*)

------------------------------------------------------------
-- CONFIG
------------------------------------------------------------

set sourceCalendarName to "Exchange Calendar"
set targetCalendarName to "iCloud Calendar"

set daysBack to 10
set daysAhead to 30

set docsPath to (POSIX path of (path to documents folder))
set syncDir to docsPath & "ExchangeSync/"
set syncFile to syncDir & "lastSync.txt"
set logFile to syncDir & "sync.log"

do shell script "mkdir -p " & quoted form of syncDir

------------------------------------------------------------
-- LOAD LAST SYNC
------------------------------------------------------------

set lastSyncDate to (current date) - (daysBack * days)
try
	set lastSyncDate to my parseISODate(do shell script "cat " & quoted form of syncFile)
end try

set nowDate to current date
set startDate to nowDate - (daysBack * days)
set endDate to nowDate + (daysAhead * days)

------------------------------------------------------------
-- STATE
------------------------------------------------------------

set newCount to 0
set updateCount to 0
set deleteCount to 0
set exchangeUIDs to {}

------------------------------------------------------------
-- MAIN SYNC
------------------------------------------------------------

try
	tell application "Calendar"
		set sourceCal to calendar sourceCalendarName
		set targetCal to calendar targetCalendarName
		
		set sourceEvents to every event of sourceCal ¬
			whose start date ≥ startDate ¬
			and start date ≤ endDate ¬
			and modification date > lastSyncDate
		
		repeat with e in sourceEvents
			
			if recurrence of e is not missing value then
				-- skip recurring
			else
				set srcUID to uid of e
				if srcUID is missing value then next
				
				set end of exchangeUIDs to srcUID
				set marker to "[EXCHANGE_UID=" & srcUID & "]"
				
				set srcSummary to my safeText(summary of e)
				set srcStart to start date of e
				set srcEnd to end date of e
				set srcLocation to my safeText(location of e)
				
				if srcSummary is "" or srcStart is missing value then next
				
				set matches to (every event of targetCal whose notes contains marker)
				
				if (count of matches) > 0 then
					set tEvent to item 1 of matches
					set summary of tEvent to srcSummary
					set start date of tEvent to srcStart
					set end date of tEvent to srcEnd
					set location of tEvent to srcLocation
					set updateCount to updateCount + 1
				else
					make new event at end of events of targetCal with properties ¬
						¬
							¬
								¬
									{summary:srcSummary, start date:srcStart, end date:srcEnd, location:srcLocation, notes:marker} ¬
										
					set newCount to newCount + 1
				end if
			end if
		end repeat
	end tell
	
	------------------------------------------------------------
	-- DELETE ORPHANED EVENTS
	------------------------------------------------------------
	
	tell application "Calendar"
		repeat with tEvent in every event of targetCal
			try
				set tNotes to notes of tEvent
				if tNotes contains "[EXCHANGE_UID=" then
					
					-- ✅ SAFE UID EXTRACTION
					set AppleScript's text item delimiters to "[EXCHANGE_UID="
					set parts to text items of tNotes
					set AppleScript's text item delimiters to "]"
					set tUID to item 1 of text items of (item 2 of parts)
					set AppleScript's text item delimiters to ""
					
					set tStart to start date of tEvent
					if tStart ≥ startDate and tStart ≤ endDate then
						if exchangeUIDs does not contain tUID then
							delete tEvent
							set deleteCount to deleteCount + 1
						end if
					end if
				end if
			end try
		end repeat
	end tell
	
	------------------------------------------------------------
	-- SAVE STATE
	------------------------------------------------------------
	
	set isoDate to do shell script "date -u '+%Y-%m-%dT%H:%M:%SZ'"
	do shell script "echo " & quoted form of isoDate & " > " & quoted form of syncFile
	
	my appendToFile(my buildLogEntry(newCount & " created, " & updateCount & " updated, " & deleteCount & " deleted"), logFile)
	
	display notification ¬
		(newCount & " new, " & updateCount & " updated, " & deleteCount & " deleted") ¬
			with title "Exchange → Apple Calendar Sync"
	
on error errMsg number errNum
	my appendToFile(my buildLogEntry("ERROR (" & errNum & "): " & errMsg), logFile)
	display notification errMsg with title "Exchange Sync Error"
end try

------------------------------------------------------------
-- UTILITIES
------------------------------------------------------------

on parseISODate(iso)
	try
		set y to text 1 thru 4 of iso
		set m to text 6 thru 7 of iso
		set d to text 9 thru 10 of iso
		set hh to text 12 thru 13 of iso
		set mm to text 15 thru 16 of iso
		set ss to text 18 thru 19 of iso
		
		set dt to current date
		set year of dt to y as integer
		set month of dt to m as integer
		set day of dt to d as integer
		set time of dt to (hh as integer) * hours + (mm as integer) * minutes + (ss as integer)
		return dt
	on error
		return current date
	end try
end parseISODate

on safeText(v)
	if v is missing value then return ""
	try
		return v as text
	on error
		return ""
	end try
end safeText

on appendToFile(t, p)
	do shell script "echo " & quoted form of t & " >> " & quoted form of p
end appendToFile

on buildLogEntry(msg)
	set ts to do shell script "date '+%Y-%m-%d %H:%M:%S'"
	return ts & " | " & msg
end buildLogEntry
