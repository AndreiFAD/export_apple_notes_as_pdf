use AppleScript version "2.4"
use scripting additions
use framework "Foundation"
use framework "AppKit"

-- classes, constants, and enums used
property NSAutoPagination : a reference to 0
property NSClipPagination : a reference to 2
property NSThread : a reference to current application's NSThread
property NSPrintJobSavingURL : a reference to current application's NSPrintJobSavingURL
property NSPrintOperation : a reference to current application's NSPrintOperation
property NSPrintSaveJob : a reference to current application's NSPrintSaveJob
property |NSURL| : a reference to current application's |NSURL|
property NSString : a reference to current application's NSString
property NSTextView : a reference to current application's NSTextView
property NSPrintInfo : a reference to current application's NSPrintInfo
property NSAttributedString : a reference to current application's NSAttributedString
property NSDictionary : a reference to current application's NSDictionary
property NSUTF8StringEncoding : a reference to current application's NSUTF8StringEncoding
property theResult : false -- whether it succeeded or not
global exportFolder

set exportFolder to "" & (path to desktop folder) & "Exported_Notes:"
tell application "Finder" to if not (exists folder exportFolder) then make new folder at desktop with properties {name:"Exported_Notes"}

-- extra lines for proper character encoding
set htmlstart to "<html><head><meta charset='utf8'></head><body>"
set htmlend to "</body></html>"

tell application "Notes"
	set attachmentLog to open for access (exportFolder & "_attachments.txt") with write permission
	repeat with theNote in notes
		-- Write the body of the note out to file as PDF
    
                -- Create a timestamp string with an extra random number to ensure files are unique!
		set theDate to (current date)
		set {year:y, month:m, day:d} to theDate
		set date_format to (y * 10000 + m * 100 + d) as string 
    
		set {hours:h, minutes:m, seconds:s} to theDate
		set pre to "AM"
		if (h > 12) then
			set h to (h - 12)
			set pre to "PM"
		end if
		set time_format to ((h & "" & m & "" & s & "" & pre) as string) & "_" & (random number from 100 to 999) as text

		set thePath to POSIX path of alias (noteNameToFilePath((name of theNote as string) & "_" & date_format & "_" & time_format ) of me)
		set thePath to (NSString's stringWithString:thePath) 

		set newPath to (thePath's stringByDeletingPathExtension()'s stringByAppendingPathExtension:"pdf")
		set aString to (NSString's stringWithString:( htmlstart & (body of theNote) as string) & htmlend )
		set aData to (aString's dataUsingEncoding:NSUTF8StringEncoding)
		set styledText to (NSAttributedString's alloc()'s initWithHTML:aData documentAttributes:(missing value))
		(my saveStyledText:styledText asPDFToFile:newPath)
		-- write attachments to text file
		if (count of (attachments of theNote)) > 0 then
			write (linefeed & name of theNote & ":" & linefeed) to attachmentLog
			repeat with theAttachment in attachments of theNote
				write ("* " & name of theAttachment & linefeed) to attachmentLog
			end repeat
		end if
	end repeat
	close access attachmentLog
end tell

-------------------------------------------------------- Handlers -----------------------------------------------------------
on replaceText(find, replace, subject)
	set prevTIDs to text item delimiters of AppleScript
	set text item delimiters of AppleScript to find
	set subject to text items of subject
	set text item delimiters of AppleScript to replace
	set subject to "" & subject
	set text item delimiters of AppleScript to prevTIDs
	return subject
end replaceText

on noteNameToFilePath(noteName)
	return (exportFolder & replaceText(":", "_", noteName) & ".html")
end noteNameToFilePath

on saveStyledText:styledText asPDFToFile:newPath
	-- create print info for saving to file
	set destURL to |NSURL|'s fileURLWithPath:newPath
	set printInfo to NSPrintInfo's alloc()'s initWithDictionary:(NSDictionary's dictionaryWithObject:destURL forKey:(NSPrintJobSavingURL)) -- sets destination
	printInfo's setJobDisposition:NSPrintSaveJob -- save to file job
	printInfo's setHorizontalPagination:NSClipPagination
	printInfo's setVerticalPagination:NSAutoPagination
	printInfo's setHorizontallyCentered:false
	printInfo's setVerticallyCentered:false
	-- get page size and margins
	set pageSize to printInfo's paperSize()
	set theLeft to printInfo's leftMargin()
	set theRight to printInfo's rightMargin()
	set theTop to printInfo's topMargin()
	-- make a very deep text view
	set theView to NSTextView's alloc()'s initWithFrame:{{0, 0}, {(pageSize's width) - theLeft - theRight, 3.0E+38}}
	theView's setHorizontallyResizable:false
	-- put in the text
	theView's textStorage()'s setAttributedString:styledText
	-- size to fit; must be done on the main thread
	if NSThread's isMainThread() then
		theView's sizeToFit()
	else
		theView's performSelectorOnMainThread:"sizeToFit" withObject:(missing value) waitUntilDone:true
	end if
	-- create print operation and run it
	set printOp to NSPrintOperation's printOperationWithView:theView printInfo:printInfo
	printOp's setShowsPrintPanel:false
	printOp's setShowsProgressPanel:false
	if NSThread's isMainThread() then
		set my theResult to printOp's runOperation()
	else
		my performSelectorOnMainThread:"runPrintOperation:" withObject:printOp waitUntilDone:true
	end if
end saveStyledText:asPDFToFile:

on runPrintOperation:printOp -- on main thread
	set my theResult to printOp's runOperation()
end runPrintOperation:
