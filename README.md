 This is the Jet Reports fork of EPPlus ( https://epplus.codeplex.com/ , (C) 2011-2017 Jan Källman and others as noted in the source code.)

### How this fork differs from the mainline:
* We enforce [a Coding Standard](https://github.com/jetreports/EPPlus/wiki/Coding-Standard), which requires:
	* Unit tests for all code changes,
	* Full-word variable names,
	* XML method comments,
	* The removal of commented-out code,
	* And more!
	    
	
* This fork contains dozens of bug fixes and enhancements:
	* Improved excel function implementations,
	* Better support for Excel-generated Charts,
	* Re-implementation of the CellStore,
	* Re-implementation of named ranges,
	* Bugfixes to the ExcelWorkbook/Worksheet and other core classes,
	* Other fixes and improvements as noted in the commit history.
* This fork is NOT the source code for the EPPlus NuGet Package; we do not maintain the NuGet Package.
	* The source code for the NuGet can be found [here](https://epplus.codeplex.com/).

### Why these changes are not on the mainline:
* There are some half-cooked features in this fork (We've added support for updating existing charts, but not inserting new charts, for example); we need these features, but Jan and swmal are understandably hesitant to include these half-cooked features in the EPPlus mainline. 
* Our coding standard requirement makes it much easier to read and maintain the code going forward, but much harder to merge with a non-standardized fork of EPPlus.
* We used to be responsible for approximately 75% of all EPPlus pull requests, so it made sense for us to maintain our own fork so we could handle our own pull requests in a timely manner. 

### About [Jet Global](https://www.jetglobal.com/):
Jet Global is one of the world leaders in business reporting and analytics, providing unparalleled access to data through fast and flexible solutions that are cost effective, provide rapid time-to-value, and are built specifically for the needs of Microsoft Dynamics ERP users. Founded in 2002, Jet Global is headquartered in the very cool city of Portland, OR. If you're looking for a sweet new gig, [we're hiring](https://www.jetglobal.com/careers/) and would love to hear from you!
