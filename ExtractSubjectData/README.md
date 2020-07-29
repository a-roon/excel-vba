This macro can be used to extract the selected subjects' data from a raw/dump file having multiple tabs for different listings.

Usage:
Press Alt+F8, go to macro module and update the following line to suite your need.

subcol = wks.Rows(1).Find("Subject", LookAt:=xlPart).Column
		
		Rows(1) = Row number of dataset header row
		"Subject" = Column label of subject ID column
