--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
--- LYNC LAUNCHER
---		Originally Created By Ben Toms, Pentland Brands Plc 27/06/12
---		Updated by Ben Toms, Pentland Brands Plc 19/09/12
---		v2 Updated by Ben Toms, Pentland Brands Plc 23/04/13	
---		v2.1 Updated by Ben Toms, Pentland Brands Plc 23/09/13, added EndPointCache delete	
--------------------------------------------------------------------------------------------------------------------------
--- SYNOPSIS
---		Queries both local & domain directories for user details, writes these details
---		to the Microsoft Lync plist & then launches Lync.
--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
--- User Information
--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------

-- Get logged in users Username
set loggedInUser to do shell script "whoami"

-- Get UniqueID for user (over 1k = domain account)
set accountType to do shell script "dscl . -read /Users/" & loggedInUser & " UniqueID | awk '{print $2}'"

-- Get Mac OS version
set OSVersion to do shell script "/usr/bin/defaults read /System/Library/CoreServices/SystemVersion ProductVersion | awk '{print substr($1,1,4)}'"

-- Get domain IP
set domainIP to do shell script "host pentland.com | head  -1 | awk '{print $4}'"

---------
---------
-- Checks to see if account is an AD Account, if it's not exits
---------
---------
if 1000 is greater than accountType then
	-- Quit Script of not an AD Account, else proceed
	tell me to quit
else
	
	---------
	---------
	-- If user is an AD account, get details
	---------
	---------
	
	-- Get AD NodeName
	set nodeName to do shell script "dscl . -read /Users/" & loggedInUser & " | awk '/^OriginalNodeName:/,/^Password:/' | head -2 | tail -1 | cut -c 2-"
	
	-- Try to get Email Address from AD lookup
	set emailAddress to do shell script "dscl . -read /Users/" & loggedInUser & " EMailAddress | awk '{print $2}'"
	
	---------
	---------
	-- Checks OS version & returns user Kerberos ID
	---------
	---------
	
	-- If pentland.com IP lookup does not bgin with 10., then assume off the nw.
	if domainIP does not start with "10." then
		-- If off nw then read users KerberosID from users plist
		set kerberosID to do shell script "defaults read /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync UserLogonName"
		-- Quit if no kerberosID in users plist
		if kerberosID is equal to " " then
			tell me to quit
		end if
		--If off nw then read users SIPAdress from users plist
		set dsclSIPAddress to do shell script "defaults read /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync KerberosID "
	else
		-- If on internal NW then get details from kerberos ticket
		if (OSVersion is equal to "10.5") or (OSVersion is equal to "10.6") then
			set kerberosID to do shell script "klist | head -2 | tail -1 | awk '{print $3}'"
		else
			set kerberosID to do shell script "klist | head -2 | tail -1 | awk '{print $2}'"
		end if
		
		--  Get Lync SIP Address from AD lookup
		--set dsclSIPAddress to do shell script "dscl " & quoted form of nodeName & " -read /Users/" & loggedInUser & " dsAttrTypeNative:msRTCSIP-PrimaryUserAddress | awk '{print $2}' | cut -c 5-"
		set dsclSIPAddress to emailAddress
		
	end if
	
	---------
	---------
	-- Quit Script User does not have a SIP Address, else proceed
	---------
	---------
	if dsclSIPAddress is equal to "" then
		tell me to quit
	else
		
		---------
		---------
		-- Correct the case & replace the characters of the returned values for Lync Plist
		---------
		---------
		set the caseSIP to change_case(dsclSIPAddress, 0)
		set the userSIP to replace_chars(caseSIP, "@", "(at)")
		set the lyncPLISTSIP to replace_chars(userSIP, ".", "(dot)")
		
		---------
		---------
		-- Write values to Lync Plist
		---------
		---------
		
		-- Delete the endpointcache file as per: http://summit7systems.com/more-o365lync-troubleshooting-the-endpoint-cache/
		do shell script "rm -rf ~/Documents/\"Microsoft User Data\"/\"Microsoft Lync Data\"/sip_" & dsclSIPAddress & "/EndpointConfiguration.cache"
		
		-- Lync will not connect behind a proxy, so disable those settings.
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync AutoConfigProxy -bool false"
		
		-- As per the above
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync UseProxyServer -bool false"
		
		-- Accept License so we're not prompted
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync acceptedSLT140 -bool true"
		
		-- Stop prompt for Conference provider
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync DoNotShowConfProviderAlert -bool true"
		
		-- Stop prompt for Presence provider
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync DoNotShowPresenceProviderAlert -bool true"
		
		-- Stop prompt for Telephone provider
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync DoNotShowTelProviderAlert -bool true"
		
		-- Set Users SIPAddress
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync KerberosID " & dsclSIPAddress
		
		-- Set User preferences
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync UserIDMRU -array '{ LogonName=\"" & loggedInUser & "\"; UserID=\"" & emailAddress & "\";}'"
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync UserLogonName " & kerberosID
		
		-- Set to Auto-Archive conversations automatically with no reminder
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync sip:" & quoted form of lyncPLISTSIP & " -dict-add ArchiveAutomatically -bool true SaveConversation -bool true ShowSaveReminder -bool false"
		
		--Standard is password authen to allow all users to perform AD lookups & get images
		do shell script "defaults write /Users/" & loggedInUser & "/Library/Preferences/com.microsoft.Lync UseKerberos -bool false"
		
		-- Open Lync
		tell application "/Applications/Microsoft Lync.app"
			activate
		end tell
	end if
	
end if

---------
---------
-- Function to change case from, http://www.macosxautomation.com/applescript/sbrt/sbrt-06.html
---------
---------
on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars

on change_case(this_text, this_case)
	if this_case is 0 then
		set the comparison_string to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		set the source_string to "abcdefghijklmnopqrstuvwxyz"
	else
		set the comparison_string to "abcdefghijklmnopqrstuvwxyz"
		set the source_string to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	end if
	set the new_text to ""
	repeat with this_char in this_text
		set x to the offset of this_char in the comparison_string
		if x is not 0 then
			set the new_text to (the new_text & character x of the source_string) as string
		else
			set the new_text to (the new_text & this_char) as string
		end if
	end repeat
	return the new_text
end change_case
