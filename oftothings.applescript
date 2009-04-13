-- -------------------------------------
-- the general run method
-- -------------------------------------
on run
	
	display dialog "Should we start? Make sure you have all task in Omni Focus marked you want to import into Thing" buttons {"OK", "Cancel"} default button 1
	
	tell application "OmniFocus"
		tell default document
			if number of document window is 0 then
				make new document window with properties {bounds:{0, 0, 1000, 500}}
			end if
		end tell
		
		tell document window 1 of front document
			set lstTrees to selected trees of content
			if (count of lstTrees) = 0 then
				try
					display dialog "Nothing selected in the right-hand panel." & return & return & "Select material to export, and try again." & return
				end try
			else
				-- Generate a TaskPaper string of the selected content
				set blnContext to (selected view mode identifier is not equal to "project")
				set lngIndent to 0
				my ExportTrees(lstTrees, lngIndent, blnContext)
				
			end if
		end tell
	end tell
end run
-- -------------------------------------
-- Walks the omni focus tree
-- -------------------------------------
on ExportTrees(lstTrees, lngIndent, blnContextView)
	-- if the tree is a task give full detail
	-- else just name and any note
	--	set strTP to ""
	set duedate to ""
	
	using terms from application "OmniFocus"
		repeat with oTree in lstTrees
			-- intialize task string
			set strTP to ""
			set notes to ""
			set tags to ""
			set oValue to value of oTree
			try
				set strName to name of oValue
			on error
				set strName to "Inbox"
			end try
			if length of strName > 0 then
				set strName to my Esc(strName)
			end if
			
			if strName ­ "Inbox" then
				set strNote to note of oValue
				if length of strNote > 0 then
					set strNote to my Esc(strNote)
				end if
			end if
			
			set clValue to class of oValue
			if (clValue is not equal to task) and (clValue is not equal to inbox task) then
				
				-- Project or Folder
				if clValue is not equal to folder then
					if clValue is not equal to project then
						--Inbox (No details)
						set strTP to strTP & "Inbox:" & return
						
					else
						-- Project (Name and possibly note)
						if length of strName > 0 then
							set strTP to strTP & strName & ":" & return
							if length of strNote > 0 then
								set notes to strNote & return
							end if
						end if
					end if
				else
					-- Folder (Just name - no note)
					set strTP to strTP & strName & ":" & return
				end if
				
			else -- Task (with details from specified columns)
				
				
				-- set recFields to {fldName:name of oValue, fldNote:note of oValue, fldDone:completed of oValue, fldContext:strContext, fldStartDate:start date of oValue, flddueDate:due date of oValue, fldDoneDate:completion date of oValue, fldDuration:estimated minutes of oValue, fldFlagged:flagged of oValue}
				
				
				-- write first line of task, followed by tags
				set lstLines to paragraphs of strName
				
				set strTP to strTP & item 1 of lstLines
				
				-- Add any tags
				set oContext to context of oValue
				if oContext is not equal to missing value then
					set tags to " @" & name of oContext & ","
				end if
				
				set dteStart to start date of oValue
				if dteStart is not equal to missing value then
					set tags to tags & " @start(" & my DateString(dteStart) & ")" & ","
				end if
				
				set dteDue to due date of oValue
				if dteDue is not equal to missing value then
					set duedate to my DateStringThings(dteDue)
				else
					set duedate to ""
				end if
				
				set lngDurn to estimated minutes of oValue
				if lngDurn is not equal to missing value then
					set tags to tags & " " & (lngDurn as string) & "min" & ","
				end if
				
				if flagged of oValue then
					set tags to tags & " @flag" & ","
				end if
				
				if completed of oValue then
					set tags to tags & " @done" & ","
				end if
				
				-- project if we know
				set aProject to containing project of oValue
				if aProject is not equal to missing value then
					set tags to tags & " @" & name of aProject & ","
				end if
				
				
				
				set strTP to strTP & return
				
				-- write any remaining lines of task as note text
				if length of lstLines > 1 then
					repeat with strLine in rest of lstLines
						set strLine to my RTrim(strLine)
						if length of strLine > 0 then
							-- change any trailling : to :-, to avoid misinterpretation as a header
							if last character of strLine is not equal to ":" then
								set notes to notes & strLine & return
							else
								set notes to notes & strLine & "-" & return
							end if
						end if
					end repeat
				end if
				
				-- append any attached note text
				set lstLines to paragraphs of strNote
				
				repeat with strLine in lstLines
					set strLine to my RTrim(strLine)
					if length of strLine > 0 then
						-- change any trailling : to :-
						if last character of strLine is not equal to ":" then
							set notes to notes & strLine & return
						else
							set notes to notes & strLine & "-" & return
						end if
					end if
				end repeat
				
			end if
			
			-- if the current node has sub-trees then recurse
			set lstSubTrees to trees of oTree
			if (count of lstSubTrees) > 0 then
				if (clValue is not equal to project) and (clValue is not equal to item) then
					set lngNewIndent to lngIndent + 1
				else
					set lngNewIndent to lngIndent
				end if
				set strTP to strTP & ExportTrees(lstSubTrees, lngNewIndent, blnContextView)
			end if
			--	my log_event(my Esc(strTP))
			my createThingTask(my Esc(strTP), my Esc(notes), tags, duedate)
		end repeat
	end using terms from
	
end ExportTrees
-- -------------------------------------
-- trims a text
-- -------------------------------------
on RTrim(someText)
	local someText
	
	repeat until someText does not end with return
		if length of someText > 1 then
			set someText to text 1 thru -2 of someText
		else
			set someText to ""
		end if
	end repeat
	
	return someText
end RTrim
-- -------------------------------------
-- converts dates into a string
-- -------------------------------------
on DateString(dte)
	-- yyyy-mm-dd hh:mm
	set strDate to ""
	if dte is not equal to missing value then
		set lngMonth to month of dte as integer
		set strMonth to lngMonth as string
		if lngMonth < 10 then set strMonth to "0" & strMonth
		
		set lngDay to day of dte as integer
		set strDay to lngDay as string
		if lngDay < 10 then set strDay to "0" & strDay
		
		set strDate to strDate & (year of dte) & "-" & strMonth & "-" & strDay
		
		set lngHrs to (hours of dte) as integer
		set lngmins to (minutes of dte) as integer
		
		if (lngHrs > 0) or (lngmins > 0) then
			set strHrs to lngHrs as string
			if lngHrs < 10 then set strHrs to "0" & strHrs
			
			set strMins to lngmins as string
			if lngmins < 10 then set strMins to "0" & strMins
			
			set strDate to strDate & " " & strHrs & ":" & strMins
		end if
	end if
	return strDate
end DateString

on DateStringThings(dte)
	-- dd-mm-yyyy
	set strDate to ""
	if dte is not equal to missing value then
		set lngMonth to month of dte as integer
		set strMonth to lngMonth as string
		if lngMonth < 10 then set strMonth to "0" & strMonth
		
		set lngDay to day of dte as integer
		set strDay to lngDay as string
		if lngDay < 10 then set strDay to "0" & strDay
		
		set strDate to strDate & strDay & "-" & strMonth & "-" & (year of dte)
		
	end if
	return strDate
end DateStringThings

-- -------------------------------------
-- escpas a string by replacing the @ symbol
-- -------------------------------------
on Esc(str)
	set str to my EscAmpersand(str, "@", "_@")
end Esc
-- -------------------------------------
-- a simple replace methhod
-- -------------------------------------
on EscAmpersand(str, pattern, replace)
	set strOldDelim to text item delimiters
	
	set text item delimiters to pattern
	set lstParts to text items of str
	set lngParts to count of lstParts
	if lngParts > 1 then
		
		set strNew to item 1 of lstParts
		repeat with n from 2 to lngParts
			set strNew to strNew & replace & item n of lstParts
		end repeat
		set text item delimiters to strOldDelim
		return strNew
	else
		set text item delimiters to strOldDelim
		return str
	end if
end EscAmpersand
-- -------------------------------------
-- A simple logging mechanism
-- -------------------------------------

on log_event(theMessage)
	set theLine to (do shell script Â
		"date  +'%Y-%m-%d %H:%M:%S'" as string) Â
		& " " & theMessage
	do shell script "echo " & "\"" & theLine & "\"" & Â
		" >> /import-events.log"
end log_event

-- -------------------------------------
-- import the task into things, we have to use key events 
-- and the clipboard since things do not have apple script support yet
-- -------------------------------------
on createThingTask(subject, notes, tags, duedate)
	activate application "Things"
	
	-- jump to inbox
	-- delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"0\" using {option down, command down}"
	my doWithTimeout(uiScript, timeoutSeconds)
	
	-- create new task
	-- delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"n\" using command down"
	my doWithTimeout(uiScript, timeoutSeconds)
	
	-- set subject
	set the clipboard to subject
	
	delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"v\" using command down"
	my doWithTimeout(uiScript, timeoutSeconds)
	
	-- jump to tags
	-- delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"	\""
	my doWithTimeout(uiScript, timeoutSeconds)
	
	set the clipboard to tags
	
	--delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"v\" using command down"
	my doWithTimeout(uiScript, timeoutSeconds)
	
	-- jump to notes
	-- delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"	\""
	my doWithTimeout(uiScript, timeoutSeconds)
	
	set the clipboard to notes
	
	-- delay 1
	set timeoutSeconds to 1.0
	set uiScript to "keystroke \"v\" using command down"
	my doWithTimeout(uiScript, timeoutSeconds)
	
	-- Send tab to jump to duedate, but only if due date is present. - glennr 13/04/09
	-- if you are like me and not using hard Due Dates in your GTD, then this is a problem 
	-- since Things (1.0.4) will automatically populate today's date if you activate this field.
	if (duedate is not equal to "None") and (duedate is not equal to "") then
		
		--tab
		delay 1
		set timeoutSeconds to 1.0
		set uiScript to "keystroke \"	 \""
		my doWithTimeout(uiScript, timeoutSeconds)
		
		delay 1
		set timeoutSeconds to 1.0
		set uiScript to "keystroke tab"
		my doWithTimeout(uiScript, timeoutSeconds)
		
		set the clipboard to duedate
		
		delay 1
		set timeoutSeconds to 1.0
		set uiScript to "keystroke \"v\" using command down"
		my doWithTimeout(uiScript, timeoutSeconds)
		
	end if
	
end createThingTask


-- -------------------------------------
-- DO SOMETHING WITH A TIMEOUT
-- -------------------------------------
on doWithTimeout(uiScript, timeoutSeconds)
	set endDate to (current date) + timeoutSeconds
	repeat
		try
			run script "tell application \"System Events\"
" & uiScript & "
end tell"
			exit repeat
		on error errorMessage
			if ((current date) > endDate) then
				error "Can not " & uiScript
			end if
		end try
	end repeat
end doWithTimeout
