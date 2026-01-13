Please develop two new ExcelDNA UDF's called  STRING_COMMON and STRING_DIFF, adding htem to the collection of existing UDFs  

Both take two strings and an integer:  

- STRING_COMMON returns a dynamic array of all common substrings of greatest possible length but at least of the length denoted by the integer.  For example: if S1="Hello there, how are you" and S2="Hello there how are you",  and N=5, two elements would be returned: "Hello there" and " how are you".  If S1="Hello!" and S2="Yelp!" and N=2, then only one element, "el" would be returned.  However, if N=1, then "el" and "!" would be returned. An empty array is returned otherwise.

- STRING_DIFF returns the longest possible substrings that are different (but with minimum length N). So with S1="Hello there, how are you" and S2="Hello there how are you", with N=2, nothing would be returned, but with N=1, "," would be returned.

Note that that the E# add-in is used, which imposes a contraint on then version of C# that can be used (I believe it's v10). Also, no Nuget libraries can be used. Please review the current set of UDFs for proven working patterns.