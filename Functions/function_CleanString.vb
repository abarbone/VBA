Function CleanString(InputString As String, Optional UnwantedCharacters As String = "<,>,:,char(34),/,\,|,?,*,$", Optional ReplaceCharacter As String = "")
    ' This function replaces unwanted characters in a text string.
    ' The unwanted characters can be a single character or a string of characters.
    ' If providing unwanted character, and not using the default characters, enter the characters as a single string with entries separated by commas.
    ' If a replacement character is not input, the unwanted characters and strings will simply be removed.
    ' Commit date: 2021-02-07

    ' Use a new variable so that the input string is not overwritten by this function.
    CleanString = InputString

    For Each UnwantedCharacter In Split(UnwantedCharacters, ",")
        CleanString = Replace(CleanString, UnwantedCharacter, ReplaceCharacter)
    Next

End Function
