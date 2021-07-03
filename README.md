# debug for OpenOffice and LibreOffice

If you ever used the BASIC programming language that is embedded in the OpenOffice/LibreOffice suite of applications, then you are likely to be disappointed with the builtin debugging facilities. To perform this type of debugging you enter command like:

* msgbox oDoc.Dbg_Properties
* msgbox oDoc.Dbg_Methods
* msgbox oDoc.Dbg_SupportedInterfaces



```
sub debug(optional oObject as object)
	' Utility routine to write an objects properties, methods and supported
	' interfaces to a text file.
	' Calls the function ShellSort()
	' Tested on Ubuntu 20.04.2 and LibreOffice Version: 6.4.7.2
	' Author: Ian Stewart. Date: July 2021
	' Usage: Add debug subroutine and ShellSort function to your Basic code.
	' In your Basic code add a line to call debug. E.g. debug(oDocument)
	dim s(2) as string '3 dimensions to string array
	dim sData as string ' temp string
	dim i as integer	
	dim j as integer
	dim iNum as integer
	dim sLine as String
	dim sMsg as String
	dim sPathFile as String
		
	if IsMissing (oObject) then
		msgbox "No object supplied. Perform testing with 'ThisComponent' object." 
		oObject = ThisComponent
	end if
	
	' Add data to the string array
	s(0) = oObject.Dbg_Properties
	s(1) = oObject.Dbg_Methods
	s(2) = oObject.Dbg_SupportedInterfaces	

	' Properties, Methods and Supported Interfaces
	for j = 0 to 2
		' Split into two parts array. Heading and Data	 
		' Heading is first two lines ends with a colon, ":"
		aTwoPart = split(s(j), ":")
    	l = LBound(aTwoPart)
    	u = UBound(aTwoPart)
		
		' Some objects have no Support Interfaces so there is no data
		if u = 0 then
			sHeading = aTwoPart(0)
			sData = ""
			goto Skip1
		end if
		
		' Make the header string as one line with newline
		sHeading = join(split(aTwoPart(0), chr(10)), " ") + ":" + chr(13)
		
		' Body has random chr(10)'s throughout. Remove:	
		sData = replace(aTwoPart(1), chr(10), " ")
	
		if j = 2 then
			' For Supported Interfaces
			' Remove groups of blank spaces
			for x = 6 to 2 step -1
				sData = replace(sData, space(x), " ")
			next x
			sData = trim(sData)
			sData = replace(sData, " ", chr(13))
			' Use shell sort function to sort the array
			aData = split(sData, Chr(13))
			aSorted_Data = ShellSort(aData)
			sData = join(aSorted_Data, chr(13))			
		end if

		if j < 2 then
			' For properties and methods
			' Create an array based on semi-colon & space, 
			' however in some cases there is semicolon and double space
			aData = split(sData, "; ")
			
	    	l = LBound(aData)
	    	u = UBound(aData)
	
			' Clean up the data one line at a time.
	    	for i = l to u	
				aData(i) = trim(aData(i))
				aData(i) = replace(aData(i), "Sbx", "")
				aData(i) = replace(aData(i), "( ", "(")
				aData(i) = replace(aData(i), " )", ")")
				aData(i) = replace(aData(i), ",", ", ")
				aData(i) = replace(aData(i), "  ", " ")	
			
				' Split the line and make lowercase bracketed field in Methods
				aLinePart = split(aData(i), " (")
				if UBound(aLinePart) = 1 then
					aLinePart(1) = LCase(aLinePart(1))
					aData(i) = join(aLinePart, " (")
				end if
			
				' Split the line and make the first field have a constant length
				aLinePart = split(aData(i), " ")
	    		' Pad first field with spaces.
	    		if len(aLinePart(0)) < 13 then
	    			aLinePart(0) = aLinePart(0) + space(12 - len(aLinePart(0)))
	    		end if
				aData(i) = join(aLinePart, " ")
						
			next i	
			
			' Use shell sort routine to sort the array
			aSorted_Data = ShellSort(aData)
		    l = LBound(aSorted_Data)
		    u = UBound(aSorted_Data)
	    					
			' Convert sorted array into a string with newlines
			sData = join(aSorted_Data(), chr(13))	
	
		end if 
		
		Skip1:
		
		' Output to a file. Needs to be checled to work with MS	
		iNum = Freefile
		' CurDir doesn't work properly. 
		' ie. CurDir + /dgb_info.txt will be HOME folder
		' To do: Test this on MS platform
		
		sPathFile = CurDir + "/dbg_info.txt"		
		
		' Output for Properties, then append for methods and supported interface
		if j = 0 then		
			open sPathFile for output As iNum
		else:
			open sPathFile for append As iNum
		end if

		print #iNum, sHeading + sData + chr(13) + chr(13)		
		close #iNum	
				
	next j
	
	' Open the file with a text editor. E.g. Pluma or Gedit
	if FileExists("/usr/bin/pluma") then
		Shell ("/usr/bin/pluma", 1, sPathFile, false)
	
	elseif FileExists("/usr/bin/gedit") then
		Shell ("/usr/bin/gedit", 1, sPathFile, false)
							
	else:
		' Or post a message
		msgbox "DGB Properties, Methods and Supported Interfaces written to:" + _
				chr(13) + chr(13) + sPathFile
	end if
				
end sub

function ShellSort(mArray)
	' Routine from: https://wiki.openoffice.org/wiki/Sorting_and_searching 
	dim n as integer, h as integer, i as integer, j as integer, t as string, _
			 Ub as integer, LB as integer
	Lb = lBound(mArray)
	Ub = uBound(mArray)	 
	' compute largest increment
	n = Ub - Lb + 1
	h = 1
	if n > 14 then
	        do while h < n
	                h = 3 * h + 1
	        loop
	        h = h \ 3
	        h = h \ 3
	end if
	do while h > 0
	' sort by insertion in increments of h
	        for i = Lb + h to Ub
	                t = mArray(i)
	                for j = i - h to Lb step -h
	                        if strComp(mArray(j), t, 0) < 1 then exit for
	                        mArray(j + h) = mArray(j)
	                next j
	                mArray(j + h) = t
	        next i
	        h = h \ 3
	loop	
	ShellSort = mArray
end function

```
