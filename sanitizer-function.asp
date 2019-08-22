<%
	'Removes harmful text from a string and cuts it to a specified length
	Function sanitize(str, cut)
		Dim remove, slice
		'String slices to remove
		remove = Array("<script", "declare @q", "http", "src=", "CC:", "TO:", "href", "url=", ".exe", "%", "(", ")", "<", ">", """", "'")
		'Trim the string to a specified number of characters
		str = Left(str, cut)
		'Cycles through the list of string slices to remove
		For Each slice In remove
			str = Replace(str, slice, "", 1, -1, 1)
		Next
		'Sets return function value			
		sanitize = str
	End Function
%>