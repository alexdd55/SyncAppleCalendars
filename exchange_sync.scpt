(*
Exchange → Apple Calendar Sync
FINAL STABLE VERSION
UID-based · Update · Delete · Exchange-safe
*)

------------------------------------------------------------
-- CONFIG
------------------------------------------------------------

set sourceCalendarName to "Kalender" -- Exchange (read-only)
set targetCalendarName to "TargetCalendar" -- Local writable calendar
set daysBack to 0
set daysAhead to 30
set dryRunMode to false -- true: log planned changes, touch nothing
set logRetentionDays to 30

set docsPath to (POSIX path of (path to documents folder))
set syncDir to docsPath & "ExchangeSync/"
set syncFile to syncDir & "lastSync.txt"
set logFile to syncDir & "sync.log"

do shell script "mkdir -p " & quoted form of syncDir

------------------------------------------------------------
-- LOG ROTATION
------------------------------------------------------------
-- sync.log is append-only and otherwise grows forever. Trim to the last
-- logRetentionDays before each run. Log lines start with "YYYY-MM-DD ...",
-- so a lexical string comparison against a cutoff date is sufficient.

try
	set rotateCmd to "if [ -f " & quoted form of logFile & " ]; then cutoff=$(date -v-" & (logRetentionDays as text) & "d '+%Y-%m-%d'); awk -v c=\"$cutoff\" 'substr($0,1,10)>=c' " & quoted form of logFile & " > " & quoted form of (logFile & ".tmp") & " && mv " & quoted form of (logFile & ".tmp") & " " & quoted form of logFile & "; fi"
	do shell script rotateCmd
end try

------------------------------------------------------------
-- WINDOW
------------------------------------------------------------

set nowDate to current date
set startDate to nowDate - (daysBack * days)
set endDate to nowDate + (daysAhead * days)

------------------------------------------------------------
-- STATE
------------------------------------------------------------

set newCount to 0
set updateCount to 0
set deleteCount to 0
set dedupCount to 0
set skippedCount to 0
set recurringSkipCount to 0
set errorCount to 0
set exchangeUIDs to {}

------------------------------------------------------------
-- MAIN SYNC
------------------------------------------------------------

try
	-- Calendar.app must be running before we can address it via AppleEvents;
	-- most sync errors historically were "-600: program is not running".
	if application "Calendar" is not running then
		launch application "Calendar"
		delay 3
	end if
	
	tell application "Calendar"
		set sourceCal to calendar sourceCalendarName
		set targetCal to calendar targetCalendarName
		
		set sourceEvents to every event of sourceCal ¬
			whose start date ≥ startDate ¬
			and start date ≤ endDate
		
		-- Snapshot the target calendar ONCE per run. Doing a fresh
		-- "whose description contains marker" query per source event costs
		-- one AppleEvent (with a Calendar.app-side TCC check per candidate)
		-- against the FULL target calendar for every single source event —
		-- O(source × target), which grinds to a multi-minute hang once the
		-- target calendar accumulates a few hundred events (it never shrinks:
		-- past events fall outside the sync window and are never revisited).
		-- One bulk fetch here turns that into O(target) total, and matching
		-- becomes an in-memory string scan (no IPC) inside the loop below.
		-- Resolving events by "event id" (their stable uid) rather than by
		-- list position also avoids the "item N of every event of calendar
		-- ... kann nicht gelesen werden" (-1728) errors seen when the target
		-- collection shifts during iteration.
		set targetUIDs to uid of every event of targetCal
		set targetDescriptions to description of every event of targetCal
		-- Calendar.app's "delete" silently no-ops on events that have a
		-- recurrence set (no error thrown, event stays put) — verified live.
		-- Both delete phases below must skip recurring events rather than
		-- pretend to have removed them.
		set targetRecurrences to recurrence of every event of targetCal
		
		repeat with e in sourceEvents
			-- Each event gets its own try so one bad/unreadable event
			-- (e.g. malformed dates, stale id) can't abort the whole run.
			try
				set srcUID to uid of e
				if srcUID is missing value then next
				if (srcUID as text) is "" then next
				
				set srcSummary to my safeText(summary of e)
				set srcStart to start date of e
				set srcEnd to end date of e
				set srcLocation to my safeText(location of e)
				set srcAllDay to allday event of e
				set srcNotes to my safeText(description of e)
				-- can be missing value (single event, not a series). NOTE: this is the
				-- master RRULE(+EXDATE) string only — Calendar.app's AppleScript
				-- dictionary has no scriptable representation of a single
				-- modified occurrence (RECURRENCE-ID override). An Exchange
				-- series with one moved/retitled instance can't be synced
				-- exactly via this API; only whole-series recurrence changes
				-- (and cancelled occurrences that show up as EXDATEs in this
				-- string) are detected.
				set srcRecurrence to recurrence of e
				
				if srcSummary is "" or srcStart is missing value then next
				
				-- Guard against malformed source events (end <= start), which
				-- Calendar.app refuses to save (-10025) and would otherwise
				-- fail identically on every future run.
				if srcEnd is missing value or srcEnd ≤ srcStart then
					set skippedCount to skippedCount + 1
					next
				end if
				
				set end of exchangeUIDs to srcUID
				set marker to "[EXCHANGE_UID=" & srcUID & "]"
				-- The marker links source→target events for matching; real
				-- meeting notes are kept after it instead of being discarded,
				-- so Exchange notes actually show up (and stay in sync) in TargetCalendar.
				set fullDescription to marker
				if srcNotes is not "" then set fullDescription to fullDescription & linefeed & linefeed & srcNotes
				
				set matchIndex to 0
				repeat with i from 1 to count of targetDescriptions
					if (item i of targetDescriptions) contains marker then
						set matchIndex to i
						exit repeat
					end if
				end repeat
				
				if matchIndex > 0 then
					set tEvent to event id (item matchIndex of targetUIDs) of targetCal
					set tDescription to item matchIndex of targetDescriptions
					
					set needsUpdate to false
					
					-- Compare summary
					if (my safeText(summary of tEvent)) is not srcSummary then set needsUpdate to true
					
					-- Compare location (handle missing/"" cleanly)
					if (my safeText(location of tEvent)) is not srcLocation then set needsUpdate to true
					
					-- Compare start/end (as text, so AppleScript doesn't compare them "weirdly")
					if ((start date of tEvent) as text) is not (srcStart as text) then set needsUpdate to true
					if ((end date of tEvent) as text) is not (srcEnd as text) then set needsUpdate to true
					
					-- Sync the all-day flag too
					if (allday event of tEvent) is not srcAllDay then set needsUpdate to true
					
					-- Sync notes too (description minus the marker)
					if tDescription is not fullDescription then set needsUpdate to true
					
					-- Recurrence rule: only touch it if the source has one (otherwise it'd keep toggling on/off)
					try
						set tRec to recurrence of tEvent
					on error
						set tRec to missing value
					end try
					
					-- optional: very conservative — only compare/set if the source has a recurrence
					if srcRecurrence is not missing value then
						try
							if (srcRecurrence as text) is not (tRec as text) then set needsUpdate to true
						on error
							-- if "as text" fails, at least set it if the target has no recurrence
							if tRec is missing value then set needsUpdate to true
						end try
					end if
					
					if needsUpdate then
						if dryRunMode then
							my appendToFile(my buildLogEntry("[DRY RUN] would update: " & srcSummary & " @ " & (srcStart as text)), logFile)
						else
							set summary of tEvent to srcSummary
							set start date of tEvent to srcStart
							set end date of tEvent to srcEnd
							set location of tEvent to srcLocation
							set description of tEvent to fullDescription
							try
								set allday event of tEvent to srcAllDay
							end try
							
							-- only set recurrence if the source has one
							if srcRecurrence is not missing value then
								try
									set recurrence of tEvent to srcRecurrence
								end try
							end if
						end if
						
						set updateCount to updateCount + 1
					end if
					
				else
					-- Create new: include recurrence if it's a series
					if dryRunMode then
						my appendToFile(my buildLogEntry("[DRY RUN] would create: " & srcSummary & " @ " & (srcStart as text)), logFile)
					else
						if srcRecurrence is missing value then
							make new event at end of events of targetCal with properties ¬
								{summary:srcSummary, start date:srcStart, end date:srcEnd, location:srcLocation, description:fullDescription, allday event:srcAllDay}
						else
							make new event at end of events of targetCal with properties ¬
								{summary:srcSummary, start date:srcStart, end date:srcEnd, location:srcLocation, description:fullDescription, allday event:srcAllDay, recurrence:srcRecurrence}
						end if
					end if
					
					set newCount to newCount + 1
				end if
				
			on error errMsg number errNum
				set errorCount to errorCount + 1
				try
					my appendToFile(my buildLogEntry("EVENT ERROR (" & errNum & "): " & errMsg), logFile)
				end try
			end try
			
		end repeat
		
	end tell
	
	------------------------------------------------------------
	-- DELETE ORPHANED EVENTS
	------------------------------------------------------------
	
	tell application "Calendar"
		-- Reuse the snapshot taken above instead of re-enumerating
		-- "every event of targetCal" (source of the -1728 stale-index
		-- errors) — resolve each candidate by its stable "event id" instead.
		repeat with i from 1 to count of targetDescriptions
			try
				set tNotes to my safeText(item i of targetDescriptions)
				
				if tNotes contains "[EXCHANGE_UID=" then
					
					-- ✅ SAFE UID EXTRACTION
					set oldDelimiters to AppleScript's text item delimiters
					try
						set AppleScript's text item delimiters to "[EXCHANGE_UID="
						set parts to text items of tNotes
						set AppleScript's text item delimiters to "]"
						set tUID to item 1 of text items of (item 2 of parts)
						set AppleScript's text item delimiters to oldDelimiters
					on error errMsg number errNum
						set AppleScript's text item delimiters to oldDelimiters
						error errMsg number errNum
					end try
					
					if exchangeUIDs does not contain tUID then
						set tEvent to event id (item i of targetUIDs) of targetCal
						set tStart to start date of tEvent
						if tStart ≥ startDate and tStart ≤ endDate then
							if (item i of targetRecurrences) is not missing value then
								set recurringSkipCount to recurringSkipCount + 1
								my appendToFile(my buildLogEntry("ORPHAN SKIP (recurring, can't delete via AppleScript) uid=" & tUID), logFile)
							else if dryRunMode then
								my appendToFile(my buildLogEntry("[DRY RUN] would delete orphan uid=" & tUID), logFile)
								set deleteCount to deleteCount + 1
							else
								delete tEvent
								set deleteCount to deleteCount + 1
							end if
						end if
					end if
				end if
			on error errMsg number errNum
				set errorCount to errorCount + 1
				try
					my appendToFile(my buildLogEntry("DELETE ERROR (" & errNum & "): " & errMsg), logFile)
				end try
			end try
		end repeat
	end tell
	
	------------------------------------------------------------
	-- DEDUPLICATE TARGET CALENDAR
	------------------------------------------------------------
	-- Invariant: "TargetCalendar" must never contain two events with the same
	-- summary + start date + end date. Runs after create/update/delete
	-- above, so it needs a fresh snapshot (the one taken earlier is now
	-- stale). Scoped to targetCal only — never touches the read-only
	-- Exchange source calendar.
	
	tell application "Calendar"
		set dedupUIDs to uid of every event of targetCal
		set dedupSummaries to summary of every event of targetCal
		set dedupStarts to start date of every event of targetCal
		set dedupEnds to end date of every event of targetCal
		set dedupDescriptions to description of every event of targetCal
		set dedupRecurrences to recurrence of every event of targetCal
		
		set dedupTotal to count of dedupUIDs
		set visited to {}
		repeat with i from 1 to dedupTotal
			set end of visited to false
		end repeat
		
		repeat with i from 1 to dedupTotal
			if not (item i of visited) then
				try
					set iSummary to my safeText(item i of dedupSummaries)
					set iStart to item i of dedupStarts
					set iEnd to item i of dedupEnds
					
					set dupIndices to {i}
					repeat with j from (i + 1) to dedupTotal
						if not (item j of visited) then
							if (my safeText(item j of dedupSummaries)) is iSummary then
								if (item j of dedupStarts) = iStart then
									if (item j of dedupEnds) = iEnd then
										set end of dupIndices to j
									end if
								end if
							end if
						end if
					end repeat
					
					if (count of dupIndices) > 1 then
						-- A shared name+start+end does NOT always mean a sync
						-- duplicate: two events tagged with DIFFERENT real
						-- Exchange UIDs are two genuinely distinct meetings
						-- that happen to coincide (verified live — e.g. two
						-- separate "Daily" standups at the same slot). Only
						-- collapse a group when at most one distinct
						-- Exchange event is represented in it.
						set taggedIndices to {}
						set distinctTags to {}
						repeat with dupIdx in dupIndices
							set idx to contents of dupIdx
							set tag to my extractExchangeUID(my safeText(item idx of dedupDescriptions))
							if tag is not "" then
								set end of taggedIndices to idx
								if distinctTags does not contain tag then set end of distinctTags to tag
							end if
						end repeat
						
						set toDelete to {}
						if (count of distinctTags) > 1 then
							-- Different real Exchange events — never delete a
							-- tagged copy. Only untagged (not Exchange-linked)
							-- local extras are safe to remove.
							repeat with dupIdx in dupIndices
								set idx to contents of dupIdx
								if taggedIndices does not contain idx then set end of toDelete to idx
							end repeat
							if (count of toDelete) < ((count of dupIndices) - 1) then
								try
									my appendToFile(my buildLogEntry("DEDUP NOTICE: \"" & iSummary & "\" @ " & (iStart as text) & " matches " & (count of distinctTags) & " distinct Exchange events — left tagged copies alone"), logFile)
								end try
							end if
						else
							-- True duplicate: 0 or 1 distinct Exchange event
							-- behind this group. Keep one copy (prefer the
							-- tagged one), delete the rest.
							set keepIndex to item 1 of dupIndices
							if (count of taggedIndices) > 0 then set keepIndex to item 1 of taggedIndices
							repeat with dupIdx in dupIndices
								set idx to contents of dupIdx
								if idx is not keepIndex then set end of toDelete to idx
							end repeat
						end if
						
						repeat with dupIdx in dupIndices
							set item (contents of dupIdx) of visited to true
						end repeat
						
						repeat with delIdx in toDelete
							set idx to contents of delIdx
							try
								if (item idx of dedupRecurrences) is not missing value then
									set recurringSkipCount to recurringSkipCount + 1
									my appendToFile(my buildLogEntry("DEDUP SKIP (recurring, can't delete via AppleScript): " & iSummary & " @ " & (iStart as text)), logFile)
								else if dryRunMode then
									my appendToFile(my buildLogEntry("[DRY RUN] would remove duplicate: " & iSummary & " @ " & (iStart as text)), logFile)
									set dedupCount to dedupCount + 1
								else
									set dupEvent to event id (item idx of dedupUIDs) of targetCal
									delete dupEvent
									set dedupCount to dedupCount + 1
								end if
							on error errMsg number errNum
								set errorCount to errorCount + 1
								try
									my appendToFile(my buildLogEntry("DEDUP ERROR (" & errNum & "): " & errMsg), logFile)
								end try
							end try
						end repeat
					else
						set item i of visited to true
					end if
				on error errMsg number errNum
					set item i of visited to true
					set errorCount to errorCount + 1
					try
						my appendToFile(my buildLogEntry("DEDUP SCAN ERROR (" & errNum & "): " & errMsg), logFile)
					end try
				end try
			end if
		end repeat
	end tell
	
	------------------------------------------------------------
	-- SAVE STATE
	------------------------------------------------------------
	
	if not dryRunMode then
		set isoDate to do shell script "date -u '+%Y-%m-%dT%H:%M:%SZ'"
		do shell script "echo " & quoted form of isoDate & " > " & quoted form of syncFile
	end if
	
	set summaryText to (newCount as text) & " new, " & (updateCount as text) & " updated, " & (deleteCount as text) & " deleted"
	set logText to (newCount as text) & " created, " & (updateCount as text) & " updated, " & (deleteCount as text) & " deleted"
	if dedupCount > 0 then set logText to logText & ", " & (dedupCount as text) & " duplicates removed"
	if skippedCount > 0 then set logText to logText & ", " & (skippedCount as text) & " skipped (bad dates)"
	if recurringSkipCount > 0 then set logText to logText & ", " & (recurringSkipCount as text) & " recurring deletions skipped (needs manual removal)"
	if errorCount > 0 then set logText to logText & ", " & (errorCount as text) & " errors"
	if dryRunMode then set logText to "[DRY RUN] " & logText
	
	my appendToFile(my buildLogEntry(logText), logFile)
	
	-- Notify only for changes that need attention (new/updated events, or
	-- errors); routine orphan-delete/dedup-only runs stay logged but silent.
	if dryRunMode then
		if newCount > 0 or updateCount > 0 or deleteCount > 0 or dedupCount > 0 or errorCount > 0 then
			display notification ("[DRY RUN] " & summaryText) with title "Exchange → Apple Calendar Sync"
		end if
	else
		if newCount > 0 or updateCount > 0 or errorCount > 0 then
			display notification summaryText with title "Exchange → Apple Calendar Sync"
		end if
	end if
	
	
on error errMsg number errNum
	my appendToFile(my buildLogEntry("ERROR (" & errNum & "): " & errMsg), logFile)
	display notification errMsg with title "Exchange Sync Error"
end try

------------------------------------------------------------
-- UTILITIES
------------------------------------------------------------

on extractExchangeUID(tNotes)
	if tNotes does not contain "[EXCHANGE_UID=" then return ""
	set oldDelimiters to AppleScript's text item delimiters
	try
		set AppleScript's text item delimiters to "[EXCHANGE_UID="
		set parts to text items of tNotes
		set AppleScript's text item delimiters to "]"
		set tUID to item 1 of text items of (item 2 of parts)
		set AppleScript's text item delimiters to oldDelimiters
		return tUID
	on error
		set AppleScript's text item delimiters to oldDelimiters
		return ""
	end try
end extractExchangeUID

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
