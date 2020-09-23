Attribute VB_Name = "Security"
'*********************************************************
'****************** Encryption/Decryption ****************
'Works for general text encryption/decryption in order to encrypt usernames
'and passwords all text with character string first and numeric
'later, is encrypted properly

Public Function Crypt(ByVal s_text As String)
    s_text = Replace(s_text, "0", "%r", , , vbBinaryCompare)
    s_text = Replace(s_text, "1", "#g", , , vbBinaryCompare)
    s_text = Replace(s_text, "2", "1$", , , vbBinaryCompare)
    s_text = Replace(s_text, "3", "j~", , , vbBinaryCompare)
    s_text = Replace(s_text, "4", "j#", , , vbBinaryCompare)
    s_text = Replace(s_text, "5", "3?", , , vbBinaryCompare)
    s_text = Replace(s_text, "6", "*t", , , vbBinaryCompare)
    s_text = Replace(s_text, "7", "u@", , , vbBinaryCompare)
    s_text = Replace(s_text, "8", "n!", , , vbBinaryCompare)
    s_text = Replace(s_text, "9", "&x", , , vbBinaryCompare)
    
    s_text = Replace(s_text, "a", "1!", , , vbBinaryCompare)
    s_text = Replace(s_text, "b", "2@", , , vbBinaryCompare)
    s_text = Replace(s_text, "c", "3#", , , vbBinaryCompare)
    s_text = Replace(s_text, "d", "4$", , , vbBinaryCompare)
    s_text = Replace(s_text, "e", "5$", , , vbBinaryCompare)
    s_text = Replace(s_text, "f", "7*", , , vbBinaryCompare)
    s_text = Replace(s_text, "g", "9#", , , vbBinaryCompare)
    s_text = Replace(s_text, "h", "0#", , , vbBinaryCompare)
    s_text = Replace(s_text, "i", "4@", , , vbBinaryCompare)
    s_text = Replace(s_text, "j", "7#", , , vbBinaryCompare)
    s_text = Replace(s_text, "k", "8^", , , vbBinaryCompare)
    s_text = Replace(s_text, "l", "0^", , , vbBinaryCompare)
    s_text = Replace(s_text, "m", "5%", , , vbBinaryCompare)
    s_text = Replace(s_text, "n", "a%", , , vbBinaryCompare)
    s_text = Replace(s_text, "o", "e$", , , vbBinaryCompare)
    s_text = Replace(s_text, "p", "f5", , , vbBinaryCompare)
    s_text = Replace(s_text, "q", "6$", , , vbBinaryCompare)
    s_text = Replace(s_text, "r", "h&", , , vbBinaryCompare)
    s_text = Replace(s_text, "s", "0.", , , vbBinaryCompare)
    s_text = Replace(s_text, "t", "e`", , , vbBinaryCompare)
    s_text = Replace(s_text, "u", "4r", , , vbBinaryCompare)
    s_text = Replace(s_text, "v", "7@", , , vbBinaryCompare)
    s_text = Replace(s_text, "w", "f^", , , vbBinaryCompare)
    s_text = Replace(s_text, "x", "t%", , , vbBinaryCompare)
    s_text = Replace(s_text, "y", "g@", , , vbBinaryCompare)
    s_text = Replace(s_text, "z", "h0", , , vbBinaryCompare)
    
    s_text = Replace(s_text, "A", ".2", , , vbBinaryCompare)
    s_text = Replace(s_text, "B", ".3", , , vbBinaryCompare)
    s_text = Replace(s_text, "C", ".4", , , vbBinaryCompare)
    s_text = Replace(s_text, "D", ".5", , , vbBinaryCompare)
    s_text = Replace(s_text, "E", ".6", , , vbBinaryCompare)
    s_text = Replace(s_text, "F", ".7", , , vbBinaryCompare)
    s_text = Replace(s_text, "G", ".8", , , vbBinaryCompare)
    s_text = Replace(s_text, "H", ".9", , , vbBinaryCompare)
    s_text = Replace(s_text, "I", ".0", , , vbBinaryCompare)
    s_text = Replace(s_text, "J", ".1", , , vbBinaryCompare)
    s_text = Replace(s_text, "K", "/3", , , vbBinaryCompare)
    s_text = Replace(s_text, "L", "/5", , , vbBinaryCompare)
    s_text = Replace(s_text, "M", "/7", , , vbBinaryCompare)
    s_text = Replace(s_text, "N", "/9", , , vbBinaryCompare)
    s_text = Replace(s_text, "O", "/1", , , vbBinaryCompare)
    s_text = Replace(s_text, "P", "/0", , , vbBinaryCompare)
    s_text = Replace(s_text, "Q", "/8", , , vbBinaryCompare)
    s_text = Replace(s_text, "R", "/6", , , vbBinaryCompare)
    s_text = Replace(s_text, "S", "/4", , , vbBinaryCompare)
    s_text = Replace(s_text, "T", "/2", , , vbBinaryCompare)
    s_text = Replace(s_text, "U", ";0", , , vbBinaryCompare)
    s_text = Replace(s_text, "V", ";2", , , vbBinaryCompare)
    s_text = Replace(s_text, "W", ";3", , , vbBinaryCompare)
    s_text = Replace(s_text, "X", ";4", , , vbBinaryCompare)
    s_text = Replace(s_text, "Y", ";6", , , vbBinaryCompare)
    s_text = Replace(s_text, "Z", ";7", , , vbBinaryCompare)
    
    Crypt = s_text
End Function

Public Function Decrypt(ByVal s_text As String)
    s_text = Replace(s_text, "1!", "a", , , vbBinaryCompare)
    s_text = Replace(s_text, "2@", "b", , , vbBinaryCompare)
    s_text = Replace(s_text, "3#", "c", , , vbBinaryCompare)
    s_text = Replace(s_text, "4$", "d", , , vbBinaryCompare)
    s_text = Replace(s_text, "5$", "e", , , vbBinaryCompare)
    s_text = Replace(s_text, "6#", "d", , , vbBinaryCompare)
    s_text = Replace(s_text, "7*", "f", , , vbBinaryCompare)
    s_text = Replace(s_text, "9#", "g", , , vbBinaryCompare)
    s_text = Replace(s_text, "0#", "h", , , vbBinaryCompare)
    s_text = Replace(s_text, "4@", "i", , , vbBinaryCompare)
    s_text = Replace(s_text, "7#", "j", , , vbBinaryCompare)
    s_text = Replace(s_text, "8^", "k", , , vbBinaryCompare)
    s_text = Replace(s_text, "0^", "l", , , vbBinaryCompare)
    s_text = Replace(s_text, "5%", "m", , , vbBinaryCompare)
    s_text = Replace(s_text, "a%", "n", , , vbBinaryCompare)
    s_text = Replace(s_text, "e$", "o", , , vbBinaryCompare)
    s_text = Replace(s_text, "f5", "p", , , vbBinaryCompare)
    s_text = Replace(s_text, "6$", "q", , , vbBinaryCompare)
    s_text = Replace(s_text, "h&", "r", , , vbBinaryCompare)
    s_text = Replace(s_text, "0.", "s", , , vbBinaryCompare)
    s_text = Replace(s_text, "e`", "t", , , vbBinaryCompare)
    s_text = Replace(s_text, "4r", "u", , , vbBinaryCompare)
    s_text = Replace(s_text, "7@", "v", , , vbBinaryCompare)
    s_text = Replace(s_text, "f^", "w", , , vbBinaryCompare)
    s_text = Replace(s_text, "t%", "x", , , vbBinaryCompare)
    s_text = Replace(s_text, "g@", "y", , , vbBinaryCompare)
    s_text = Replace(s_text, "h0", "z", , , vbBinaryCompare)
    
    s_text = Replace(s_text, ".2", "A", , , vbBinaryCompare)
    s_text = Replace(s_text, ".3", "B", , , vbBinaryCompare)
    s_text = Replace(s_text, ".4", "C", , , vbBinaryCompare)
    s_text = Replace(s_text, ".5", "D", , , vbBinaryCompare)
    s_text = Replace(s_text, ".6", "E", , , vbBinaryCompare)
    s_text = Replace(s_text, ".7", "F", , , vbBinaryCompare)
    s_text = Replace(s_text, ".8", "G", , , vbBinaryCompare)
    s_text = Replace(s_text, ".9", "H", , , vbBinaryCompare)
    s_text = Replace(s_text, ".0", "I", , , vbBinaryCompare)
    s_text = Replace(s_text, ".1", "J", , , vbBinaryCompare)
    s_text = Replace(s_text, "/3", "K", , , vbBinaryCompare)
    s_text = Replace(s_text, "/5", "L", , , vbBinaryCompare)
    s_text = Replace(s_text, "/7", "M", , , vbBinaryCompare)
    s_text = Replace(s_text, "/9", "N", , , vbBinaryCompare)
    s_text = Replace(s_text, "/1", "O", , , vbBinaryCompare)
    s_text = Replace(s_text, "/0", "P", , , vbBinaryCompare)
    s_text = Replace(s_text, "/8", "Q", , , vbBinaryCompare)
    s_text = Replace(s_text, "/6", "R", , , vbBinaryCompare)
    s_text = Replace(s_text, "/4", "S", , , vbBinaryCompare)
    s_text = Replace(s_text, "/2", "T", , , vbBinaryCompare)
    s_text = Replace(s_text, ";0", "U", , , vbBinaryCompare)
    s_text = Replace(s_text, ";2", "V", , , vbBinaryCompare)
    s_text = Replace(s_text, ";3", "W", , , vbBinaryCompare)
    s_text = Replace(s_text, ";4", "X", , , vbBinaryCompare)
    s_text = Replace(s_text, ";6", "Y", , , vbBinaryCompare)
    s_text = Replace(s_text, ";7", "Z", , , vbBinaryCompare)
    
    
    s_text = Replace(s_text, "%r", "0", , , vbBinaryCompare)
    s_text = Replace(s_text, "#g", "1", , , vbBinaryCompare)
    s_text = Replace(s_text, "1$", "2", , , vbBinaryCompare)
    s_text = Replace(s_text, "j~", "3", , , vbBinaryCompare)
    s_text = Replace(s_text, "j#", "4", , , vbBinaryCompare)
    s_text = Replace(s_text, "3?", "5", , , vbBinaryCompare)
    s_text = Replace(s_text, "*t", "6", , , vbBinaryCompare)
    s_text = Replace(s_text, "u@", "7", , , vbBinaryCompare)
    s_text = Replace(s_text, "n!", "8", , , vbBinaryCompare)
    s_text = Replace(s_text, "&x", "9", , , vbBinaryCompare)
    Decrypt = s_text
End Function

