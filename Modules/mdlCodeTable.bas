Attribute VB_Name = "mdlCodeTable"
Option Explicit

Public TelexArr() As Variant, VniArr() As Variant, ViqrArr() As Variant


Public UNICODE_PRECOMPOSED_TABLE() As Variant, _
        BKHCM1_TABLE() As Variant, _
        BKHCM2_TABLE() As Variant, _
        TCVN3_TABLE() As Variant, _
        UTF8_TABLE() As Variant, _
        VIETWARE_F_TABLE() As Variant, _
        VIQR_TABLE() As Variant, _
        VISCII_TABLE() As Variant, _
        VNCP_1258_TABLE() As Variant, _
        VNI_WINDOWS_TABLE() As Variant, _
        VPS_TABLE() As Variant, _
        UNICODE_COMPOSED_TABLE() As Variant



Private CodeTableMaked As Boolean



Public Enum CODE_TABLE

    UNICODE_PRECOMPOSED_TABLE_ENUM = 1          '   1
    '--------------------------
    BKHCM1_TABLE_ENUM = 3                       '   2
    BKHCM2_TABLE_ENUM = 4                       '   3
    TCVN3_TABLE_ENUM = 5                        '   4
    UTF8_TABLE_ENUM = 6                         '   5
    '--------------------------
    VIETWARE_F_TABLE_ENUM = 8                   '   6
    VIQR_TABLE_ENUM = 9                         '   7
    VISCII_TABLE_ENUM = 10                      '   8
    VNCP_1258_TABLE_ENUM = 11                   '   9
    VNI_WINDOWS_TABLE_ENUM = 12                 '   10
    VPS_TABLE_ENUM = 13                         '   11
    '--------------------------
    UNICODE_COMPOSED_TABLE_ENUM = 15            '   12
    
End Enum


Public Sub MakeCodeTable()

    CodeTableMaked = True
    
    '   =================== TELEX, VNI, VIQR CODE TABLE ==================
    
    TelexArr = Array("af", "as", "ar", "ax", "aj", "aa", "aaf", "aas", "aar", "aax", "aaj", "aw", "awf", "aws", "awr", "awx", "awj", "dd", "ef", "es", "er", "ex", "ej", "ee", "eef", "ees", "eer", "eex", "eej", "if", "is", "ir", "ix", "ij", "of", "os", "or", "ox", "oj", "oo", "oof", "oos", "oor", "oox", "ooj", "ow", "owf", "ows", "owr", "owx", "owj", "uf", "us", "ur", "ux", "uj", "uw", "uwf", "uws", "uwr", "uwx", "uwj", "yf", "ys", "yr", "yx", "yj", "AF", "AS", "AR", "AX", "AJ", "AA", "AAF", "AAS", "AAR", "AAX", "AAJ", "AW", "AWF", "AWS", "AWR", "AWX", "AWJ", "DD", "EF", "ES", "ER", "EX", "EJ", "EE", "EEF", "EES", "EER", "EEX", "EEJ", "IF", "IS", "IR", "IX", "IJ", "OF", "OS", "OR", "OX", "OJ", "OO", "OOF", "OOS", "OOR", "OOX", "OOJ", "OW", "OWF", "OWS", "OWR", "OWX", "OWJ", "UF", "US", "UR", "UX", "UJ", "UW", "UWF", "UWS", "UWR", "UWX", "UWJ", "YF", "YS", "YR", "YX", "YJ")
    VniArr = Array("a2", "a1", "a3", "a4", "a5", "a6", "a62", "a61", "a63", "a64", "a65", "a8", "a82", "a81", "a83", "a84", "a85", "d9", "e2", "e1", "e3", "e4", "e5", "e6", "e62", "e61", "e63", "e64", "e65", "i2", "i1", "i3", "i4", "i5", "o2", "o1", "o3", "o4", "o5", "o6", "o62", "o61", "o63", "o64", "o65", "o7", "o72", "o71", "o73", "o74", "o75", "u2", "u1", "u3", "u4", "u5", "u7", "u72", "u71", "u73", "u74", "u75", "y2", "y1", "y3", "y4", "y5", "A2", "A1", "A3", "A4", "A5", "A6", "A62", "A61", "A63", "A64", "A65", "A8", "A82", "A81", "A83", "A84", "A85", "D9", "E2", "E1", "E3", "E4", "E5", "E6", "E62", "E61", "E63", "E64", "E65", "I2", "I1", "I3", "I4", "I5", "O2", "O1", "O3", "O4", "O5", "O6", "O62", "O61", "O63", "O64", "O65", "O7", "O72", "O71", "O73", "O74", "O75", "U2", "U1", "U3", "U4", "U5", "U7", "U72", "U71", "U73", "U74", "U75", "Y2", "Y1", "Y3", "Y4", "Y5")
    ViqrArr = Array("a`", "a'", "a?", "a~", "a.", "a^", "a^`", "a^'", "a^?", "a^~", "a^.", "a(", "a(`", "a('", "a(?", "a(~", "a(.", "dd", "e`", "e'", "e?", "e~", "e.", "e^", "e^`", "e^'", "e^?", "e^~", "e^.", "i`", "i'", "i?", "i~", "i.", "o`", "o'", "o?", "o~", "o.", "o^", "o^`", "o^'", "o^?", "o^~", "o^.", "o+", "o+`", "o+'", "o+?", "o+~", "o+.", "u`", "u'", "u?", "u~", "u.", "u+", "u+`", "u+'", "u+?", "u+~", "u+.", "y`", "y'", "y?", "y~", "y.", " A`", "A'", "A?", "A~", "A.", "A^", "A^`", "A^'", "A^?", "A^~", "A^.", "A(", "A(`", "A('", "A(?", "A(~", "A(.", "DD", " E`", "E'", "E?", "E~", "E.", "E^", "E^`", "E^'", "E^?", "E^~", "E^.", "I`", "I'", "I?", "I~", "I.", "O`", "O'", "O?", "O~", "O.", "O^", "O^`", "O^'", "O^?", "O^~", "O^.", "O+", "O+`", "O+'", "O+?", "O+~", "O+.", "U`", "U'", "U?", "U~", "U.", "U+", "U+`", "U+'", "U+?", "U+~", "U+.", "Y`", "Y'", "Y?", "Y~", "Y.")
    
    
    '   ========================= UNICODE CODE TABLE ===================
    
    UNICODE_PRECOMPOSED_TABLE = Array( _
        ChrW$(&HE0), ChrW$(&HE1), ChrW$(&H1EA3), ChrW$(&HE3), ChrW$(&H1EA1), ChrW$(&HE2), ChrW$(&H1EA7), ChrW$(&H1EA5), _
        ChrW$(&H1EA9), ChrW$(&H1EAB), ChrW$(&H1EAD), ChrW$(&H103), ChrW$(&H1EB1), ChrW$(&H1EAF), ChrW$(&H1EB3), _
        ChrW$(&H1EB5), ChrW$(&H1EB7), ChrW$(&H111), ChrW$(&HE8), ChrW$(&HE9), ChrW$(&H1EBB), ChrW$(&H1EBD), _
        ChrW$(&H1EB9), ChrW$(&HEA), ChrW$(&H1EC1), ChrW$(&H1EBF), ChrW$(&H1EC3), ChrW$(&H1EC5), ChrW$(&H1EC7), _
        ChrW$(&HEC), ChrW$(&HED), ChrW$(&H1EC9), ChrW$(&H129), ChrW$(&H1ECB), ChrW$(&HF2), ChrW$(&HF3), ChrW$(&H1ECF), _
        ChrW$(&HF5), ChrW$(&H1ECD), ChrW$(&HF4), ChrW$(&H1ED3), ChrW$(&H1ED1), ChrW$(&H1ED5), ChrW$(&H1ED7), _
        ChrW$(&H1ED9), ChrW$(&H1A1), ChrW$(&H1EDD), ChrW$(&H1EDB), ChrW$(&H1EDF), ChrW$(&H1EE1), ChrW$(&H1EE3), _
        ChrW$(&HF9), ChrW$(&HFA), ChrW$(&H1EE7), ChrW$(&H169), ChrW$(&H1EE5), ChrW$(&H1B0), ChrW$(&H1EEB), _
        ChrW$(&H1EE9), ChrW$(&H1EED), ChrW$(&H1EEF), ChrW$(&H1EF1), ChrW$(&H1EF3), ChrW$(&HFD), ChrW$(&H1EF7), _
        ChrW$(&H1EF9), ChrW$(&H1EF5), ChrW$(&HC0), ChrW$(&HC1), ChrW$(&H1EA2), ChrW$(&HC3), ChrW$(&H1EA0), _
        ChrW$(&HC2), ChrW$(&H1EA6), ChrW$(&H1EA4), ChrW$(&H1EA8), ChrW$(&H1EAA), ChrW$(&H1EAC), ChrW$(&H102), _
        ChrW$(&H1EB0), ChrW$(&H1EAE), ChrW$(&H1EB2), ChrW$(&H1EB4), ChrW$(&H1EB6), ChrW$(&H110), ChrW$(&HC8), _
        ChrW$(&HC9), ChrW$(&H1EBA), ChrW$(&H1EBC), ChrW$(&H1EB8), ChrW$(&HCA), ChrW$(&H1EC0), ChrW$(&H1EBE), _
        ChrW$(&H1EC2), ChrW$(&H1EC4), ChrW$(&H1EC6), ChrW$(&HCC), ChrW$(&HCD), ChrW$(&H1EC8), ChrW$(&H128), _
        ChrW$(&H1ECA), ChrW$(&HD2), ChrW$(&HD3), ChrW$(&H1ECE), ChrW$(&HD5), ChrW$(&H1ECC), ChrW$(&HD4), _
        ChrW$(&H1ED2), ChrW$(&H1ED0), ChrW$(&H1ED4), ChrW$(&H1ED6), ChrW$(&H1ED8), ChrW$(&H1A0), ChrW$(&H1EDC), _
        ChrW$(&H1EDA), ChrW$(&H1EDE), ChrW$(&H1EE0), ChrW$(&H1EE2), ChrW$(&HD9), ChrW$(&HDA), ChrW$(&H1EE6), _
        ChrW$(&H168), ChrW$(&H1EE4), ChrW$(&H1AF), ChrW$(&H1EEA), ChrW$(&H1EE8), ChrW$(&H1EEC), ChrW$(&H1EEE), _
        ChrW$(&H1EF0), ChrW$(&H1EF2), ChrW$(&HDD), ChrW$(&H1EF6), ChrW$(&H1EF8), ChrW$(&H1EF4))


    VIQR_TABLE = Array( _
        "a`", "a'", "a?", "a~", "a.", "a^", "a^`", "a^'", "a^?", "a^~", "a^.", _
        "a(", "a(`", "a('", "a(?", "a(~", "a(.", "dd", " e`", "e'", "e?", "e~", _
        "e.", "e^", "e^`", "e^'", "e^?", "e^~", "e^.", "i`", "i'", "i?", "i~", _
        "i.", "o`", "o'", "o?", "o~", "o.", "o^", "o^`", "o^'", "o^?", "o^~", _
        "o^.", "o+", "o+`", "o+'", "o+?", "o+~", "o+.", "u`", "u'", "u?", "u~", _
        "u.", "u+", "u+`", "u+'", "u+?", "u+~", "u+.", "y`", "y'", "y?", "y~", "y.", _
        "A`", "A'", "A?", "A~", "A.", "A^", "A^`", "A^'", "A^?", "A^~", "A^.", _
        "A(", "A(`", "A('", "A(?", "A(~", "A(.", "DD", " E`", "E'", "E?", "E~", _
        "E.", "E^", "E^`", "E^'", "E^?", "E^~", "E^.", "I`", "I'", "I?", "I~", _
        "I.", "O`", "O'", "O?", "O~", "O.", "O^", "O^`", "O^'", "O^?", "O^~", _
        "O^.", "O+", "O+`", "O+'", "O+?", "O+~", "O+.", "U`", "U'", "U?", "U~", _
        "U.", "U+", "U+`", "U+'", "U+?", "U+~", "U+.", "Y`", "Y'", "Y?", "Y~", "Y.")


    TCVN3_TABLE = Array( _
            "µ", "∏", "∂", "∑", "π", "©", "«", " ", "»", "…", "À", "®", "ª", _
            "æ", "º", "Ω", "∆", "Æ", "Ã", "–", "Œ", "œ", "—", "™", "“", "’", _
            "”", "‘", "÷", "◊", "›", "ÿ", "‹", "ﬁ", "ﬂ", "„", "·", "‚", "‰", _
            "´", "Â", "Ë", "Ê", "Á", "È", "¨", "Í", "Ì", "Î", "Ï", "Ó", "Ô", _
            "Û", "Ò", "Ú", "Ù", "≠", "ı", "¯", "ˆ", "˜", "˘", "˙", "˝", "˚", _
            "¸", "˛", "µ", "∏", "∂", "∑", "π", "¢", "«", " ", "»", "…", "À", _
            "°", "ª", "æ", "º", "Ω", "∆", "ß", "Ã", "–", "Œ", "œ", "—", "£", _
            "“", "’", "”", "‘", "÷", "◊", "›", "ÿ", "‹", "ﬁ", "ﬂ", "„", "·", _
            "‚", "‰", "§", "Â", "Ë", "Ê", "Á", "È", "•", "Í", "Ì", "Î", "Ï", _
            "Ó", "Ô", "Û", "Ò", "Ú", "Ù", "¶", "ı", "¯", "ˆ", "˜", "˘", "˙", _
            "˝", "˚", "¸", "˛")
                                               
                                               
    VNI_WINDOWS_TABLE = Array( _
            "a¯", "a˘", "a˚", "aı", "aÔ", "a‚", "a‡", "a·", "aÂ", "a„", "a‰", _
            "aÍ", "aË", "aÈ", "a˙", "a¸", "aÎ", "Ò", "e¯", "e˘", "e˚", "eı", _
            "eÔ", "e‚", "e‡", "e·", "eÂ", "e„", "e‰", "Ï", "Ì", "Ê", "Û", "Ú", _
            "o¯", "o˘", "o˚", "oı", "oÔ", "o‚", "o‡", "o·", "oÂ", "o„", "o‰", _
            "Ù", "Ù¯", "Ù˘", "Ù˚", "Ùı", "ÙÔ", "u¯", "u˘", "u˚", "uı", "uÔ", _
            "ˆ", "ˆ¯", "ˆ˘", "ˆ˚", "ˆı", "ˆÔ", "y¯", "y˘", "y˚", "yı", "Ó", _
            "Aÿ", "AŸ", "A€", "A’", "Aœ", "A¬", "A¿", "A¡", "A≈", "A√", "Aƒ", _
            "A ", "A»", "A…", "A⁄", "A‹", "AÀ", "—", "Eÿ", "EŸ", "E€", "E’", _
            "Eœ", "E¬", "E¿", "E¡", "E≈", "E√", "Eƒ", "Ã", "Õ", "∆", "”", "“", _
            "Oÿ", "OŸ", "O€", "O’", "Oœ", "O¬", "O¿", "O¡", "O≈", "O√", "Oƒ", _
            "‘", "‘ÿ", "‘Ÿ", "‘€", "‘’", "‘œ", "Uÿ", "UŸ", "U€", "U’", "Uœ", _
            "÷", "÷ÿ", "÷Ÿ", "÷€", "÷’", "÷œ", "Yÿ", "YŸ", "Y€", "Y’", "Œ")
            
            
    VNCP_1258_TABLE = Array( _
            "aÃ", "aÏ", "a“", "aﬁ", "aÚ", "‚", "‚Ã", "‚Ï", "‚“", "‚ﬁ", "‚Ú", _
            "„", "„Ã", "„Ï", "„“", "„ﬁ", "„Ú", "", "eÃ", "eÏ", "e“", "eﬁ", _
            "eÚ", "Í", "ÍÃ", "ÍÏ", "Í“", "Íﬁ", "ÍÚ", "iÃ", "iÏ", "i“", "iﬁ", _
            "iÚ", "oÃ", "oÏ", "o“", "oﬁ", "oÚ", "Ù", "ÙÃ", "ÙÏ", "Ù“", "Ùﬁ", _
            "ÙÚ", "ı", "ıÃ", "ıÏ", "ı“", "ıﬁ", "ıÚ", "uÃ", "uÏ", "u“", "uﬁ", _
            "uÚ", "˝", "˝Ã", "˝Ï", "˝“", "˝ﬁ", "˝Ú", "yÃ", "yÏ", "y“", "yﬁ", _
            "yÚ", "AÃ", "AÏ", "A“", "Aﬁ", "AÚ", "¬", "¬Ã", "¬Ï", "¬“", "¬ﬁ", _
            "¬Ú", "√", "√Ã", "√Ï", "√“", "√ﬁ", "√Ú", "–", "EÃ", "EÏ", "E“", _
            "Eﬁ", "EÚ", " ", " Ã", " Ï", " “", " ﬁ", " Ú", "IÃ", "IÏ", "I“", _
            "Iﬁ", "IÚ", "OÃ", "OÏ", "O“", "Oﬁ", "OÚ", "‘", "‘Ã", "‘Ï", "‘“", _
            "‘ﬁ", "‘Ú", "’", "’Ã", "’Ï", "’“", "’ﬁ", "’Ú", "UÃ", "UÏ", "U“", _
            "Uﬁ", "UÚ", "›", "›Ã", "›Ï", "›“", "›ﬁ", "›Ú", "YÃ", "YÏ", "Y“", _
            "Yﬁ", "YÚ")
            
            
    UNICODE_COMPOSED_TABLE = Array( _
            "aÃ", "aÏ", "a“", "aﬁ", "aÚ", "‚", "‚Ã", "‚Ï", "‚“", "‚ﬁ", "‚Ú", _
            "„", "„Ã", "„Ï", "„“", "„ﬁ", "„Ú", "", "eÃ", "eÏ", "e“", "eﬁ", _
            "eÚ", "Í", "ÍÃ", "ÍÏ", "Í“", "Íﬁ", "ÍÚ", "iÃ", "iÏ", "i“", "iﬁ", _
            "iÚ", "oÃ", "oÏ", "o“", "oﬁ", "oÚ", "Ù", "ÙÃ", "ÙÏ", "Ù“", "Ùﬁ", _
            "ÙÚ", "ı", "ıÃ", "ıÏ", "ı“", "ıﬁ", "ıÚ", "uÃ", "uÏ", "u“", "uﬁ", _
            "uÚ", "˝", "˝Ã", "˝Ï", "˝“", "˝ﬁ", "˝Ú", "yÃ", "yÏ", "y“", "yﬁ", _
            "yÚ", "AÃ", "AÏ", "A“", "Aﬁ", "AÚ", "¬", "¬Ã", "¬Ï", "¬“", "¬ﬁ", _
            "¬Ú", "√", "√Ã", "√Ï", "√“", "√ﬁ", "√Ú", "–", "EÃ", "EÏ", "E“", _
            "Eﬁ", "EÚ", " ", " Ã", " Ï", " “", " ﬁ", " Ú", "IÃ", "IÏ", "I“", _
            "Iﬁ", "IÚ", "OÃ", "OÏ", "O“", "Oﬁ", "OÚ", "‘", "‘Ã", "‘Ï", "‘“", _
            "‘ﬁ", "‘Ú", "’", "’Ã", "’Ï", "’“", "’ﬁ", "’Ú", "UÃ", "UÏ", "U“", _
            "Uﬁ", "UÚ", "›", "›Ã", "›Ï", "›“", "›ﬁ", "›Ú", "YÃ", "YÏ", "Y“", _
            "Yﬁ", "YÚ")
            
            
    UTF8_TABLE = Array( _
            "√†", "√°", "·∫£", "√£", "·∫°", "√¢", "·∫ß", "·∫•", "·∫©", "·∫´", _
            "·∫≠", "ƒÉ", "·∫±", "·∫Ø", "·∫≥", "·∫µ", "·∫∑", "ƒë", "√®", "√©", _
            "·∫ª", "·∫Ω", "·∫π", "√™", "·ªÅ", "·∫ø", "·ªÉ", "·ªÖ", "·ªá", "√¨", _
            "√≠", "·ªâ", "ƒ©", "·ªã", "√≤", "√≥", "·ªè", "√µ", "·ªç", "√¥", _
            "·ªì", "·ªë", "·ªï", "·ªó", "·ªô", "∆°", "·ªù", "·ªõ", "·ªü", _
            "·ª°", "·ª£", "√π", "√∫", "·ªß", "≈©", "·ª•", "∆∞", "·ª´", "·ª©", _
            "·ª≠", "·ªØ", "·ª±", "·ª≥", "√Ω", "·ª∑", "·ªπ", "·ªµ", "√Ä", "√Å", _
            "·∫¢", "√É", "·∫†", "√Ç", "·∫¶", "·∫§", "·∫®", "·∫™", "·∫¨", "ƒÇ", _
            "·∫∞", "·∫Æ", "·∫≤", "·∫¥", "·∫∂", "ƒê", "√à", "√â", "·∫∫", "·∫º", _
            "·∫∏", "√ä", "·ªÄ", "·∫æ", "·ªÇ", "·ªÑ", "·ªÜ", "√å", "√ç", "·ªà", _
            "ƒ®", "·ªä", "√í", "√ì", "·ªé", "√ï", "·ªå", "√î", "·ªí", "·ªê", _
            "·ªî", "·ªñ", "·ªò", "∆†", "·ªú", "·ªö", "·ªû", "·ª†", "·ª¢", "√ô", _
            "√ö", "·ª¶", "≈®", "·ª§", "∆Ø", "·ª™", "·ª®", "·ª¨", "·ªÆ", "·ª∞", _
            "·ª≤", "√ù", "·ª∂", "·ª∏", "·ª¥")
            
                
    VISCII_TABLE = Array( _
            "‡", "·", "‰", "„", "’", "‚", "•", "§", "¶", "Á", "ß", "Â", "¢", _
            "°", "∆", "«", "£", "", "Ë", "È", "Î", "®", "©", "Í", "´", "™", _
            "¨", "≠", "Æ", "Ï", "Ì", "Ô", "Ó", "∏", "Ú", "Û", "ˆ", "ı", "˜", _
            "Ù", "∞", "Ø", "±", "≤", "µ", "Ω", "∂", "æ", "∑", "ﬁ", "˛", "˘", _
            "˙", "¸", "˚", "¯", "ﬂ", "◊", "—", "ÿ", "Ê", "Ò", "œ", "˝", "÷", _
            "€", "‹", "¿", "¡", "ƒ", "√", "Ä", "¬", "Ö", "Ñ", "Ü", "Á", "á", _
            "≈", "Ç", "Å", "∆", "«", "É", "–", "»", "…", "À", "à", "â", " ", _
            "ã", "ä", "å", "ç", "é", "Ã", "Õ", "õ", "Œ", "ò", "“", "”", "ô", _
            "ı", "ö", "‘", "ê", "è", "ë", "í", "ì", "¥", "ñ", "ï", "ó", "≥", _
            "î", "Ÿ", "⁄", "ú", "ù", "û", "ø", "ª", "∫", "º", "ˇ", "π", "ü", _
            "›", "÷", "€", "‹")
            
            
    VPS_TABLE = Array( _
            "‡", "·", "‰", "„", "Â", "‚", "¿", "√", "ƒ", "≈", "∆", "Ê", "¢", _
            "°", "£", "§", "•", "«", "Ë", "È", "»", "Î", "À", "Í", "ä", "â", _
            "ã", "Õ", "å", "Ï", "Ì", "Ã", "Ô", "Œ", "Ú", "Û", "’", "ı", "Ü", _
            "Ù", "“", "”", "∞", "á", "∂", "÷", "©", "ß", "™", "´", "Æ", "˘", _
            "˙", "˚", "€", "¯", "‹", "ÿ", "Ÿ", "∫", "ª", "ø", "ˇ", "ö", "õ", _
            "œ", "ú", "Ä", "¡", "Å", "Ç", "Â", "¬", "Ñ", "É", "Ö", "≈", "∆", _
            "à", "é", "ç", "è", "", "•", "Ò", "◊", "…", "ﬁ", "˛", "À", " ", _
            "ì", "ê", "î", "ï", "å", "µ", "¥", "∑", "∏", "Œ", "º", "π", "Ω", _
            "æ", "Ü", "‘", "ó", "ñ", "ò", "ô", "∂", "˜", "û", "ù", "ü", "¶", _
            "Æ", "®", "⁄", "—", "¨", "¯", "–", "Ø", "≠", "±", "ª", "ø", "≤", _
            "›", "˝", "≥", "ú")
            
                
    BKHCM1_TABLE = Array( _
            "ø", "æ", "¿", "¡", "¬", "›", "ﬂ", "ﬁ", "‡", "·", "‚", "◊", "Ÿ", _
            "ÿ", "⁄", "€", "}", "Ω", "ƒ", "√", "≈", "∆", "«", "„", "Â", "‰", _
            "Ê", "Á", "Ë", "…", "»", " ", "À", "Ã", "Œ", "Õ", "œ", "–", "—", _
            "È", "Î", "Í", "Ï", "Ì", "Ó", "Ô", "Ò", "", "Ú", "Û", "Ù", "”", _
            "“", "‘", "’", "÷", "ı", "˜", "ˆ", "¯", "˘", "˙", "¸", "˚", "˝", _
            "˛", "ˇ", "Å", "Ä", "Ç", "É", "Ñ", "ü", "°", "~", "¢", "£", "§", _
            "ô", "õ", "ö", "ú", "ù", "ò", "}", "Ü", "Ö", "á", "à", "â", "•", _
            "ß", "¶", "®", "©", "™", "ã", "ä", "å", "ç", "é", "ê", "è", "ë", _
            "í", "ì", "´", "≠", "¨", "Æ", "Ø", "∞", "±", "≥", "≤", "¥", "µ", _
            "∂", "ï", "î", "ñ", "ó", "ò", "∑", "π", "∏", "∫", "ª", "º", "^", _
            "{", "`", "|", "é")
            
            
    BKHCM2_TABLE = Array( _
            "a‚", "a·", "a„", "a‰", "aÂ", "Í", "ÍÏ", "ÍÎ", "ÍÌ", "ÍÓ", "ÍÂ", _
            "˘", "˘Á", "˘Ê", "˘Ë", "˘È", "˘Â", "‡", "e‚", "e·", "e„", "e‰", _
            "eÂ", "Ô", "ÔÏ", "ÔÎ", "ÔÌ", "ÔÓ", "ÔÂ", "Ú", "Ò", "Û", "Ù", "ı", _
            "o‚", "o·", "o„", "o‰", "oÂ", "ˆ", "ˆÏ", "ˆÎ", "ˆÌ", "ˆÓ", "ˆÂ", _
            "˙", "˙‚", "˙·", "˙„", "˙‰", "˙Â", "u‚", "u·", "u„", "u‰", "uÂ", _
            "˚", "˚‚", "˚·", "˚„", "˚‰", "˚Â", "y‚", "y·", "y„", "y‰", "yÂ", _
            "A¬", "A¡", "A√", "Aƒ", "A≈", " ", " Ã", " À", " Õ", " Œ", " ≈", _
            "Ÿ", "Ÿ«", "Ÿ∆", "Ÿ»", "Ÿ…", "Ÿ≈", "¿", "E¬", "E¡", "E√", "Eƒ", _
            "E≈", "œ", "œÃ", "œÀ", "œÕ", "œŒ", "œÂ", "“", "—", "”", "‘", "’", _
            "O¬", "O¡", "O√", "Oƒ", "O≈", "÷", "÷Ã", "÷À", "÷Õ", "÷Œ", "÷≈", _
            "⁄", "⁄¬", "⁄¡", "⁄√", "⁄ƒ", "⁄≈", "U¬", "U¡", "U√", "Uƒ", "U≈", _
            "€", "€¬", "€¡", "€√", "€ƒ", "€≈", "Y¬", "Y¡", "Y√", "Yƒ", "Y≈")
            
            
    VIETWARE_F_TABLE = Array( _
            "™", "¿", "∂", "∫", "¡", "°", "«", " ", "»", "…", "À", "ü", "¬", _
            "≈", "√", "ƒ", "∆", "¢", "Ã", "œ", "Õ", "Œ", "—", "£", "“", "’", _
            "”", "‘", "÷", "ÿ", "€", "Ÿ", "⁄", "‹", "ﬂ", "‚", "‡", "·", "„", _
            "§", "‰", "Á", "Â", "Ê", "Ë", "•", "È", "Ï", "Í", "Î", "Ì", "Ó", _
            "Ú", "Ô", "Ò", "Û", "ß", "Ù", "˜", "ı", "ˆ", "¯", "˘", "¸", "˙", _
            "˚", "ˇ", "™", "¿", "∂", "∫", "¡", "ó", "«", " ", "»", "…", "À", _
            "ñ", "¬", "≈", "√", "ƒ", "∆", "ò", "Ã", "œ", "Õ", "Œ", "—", "ô", _
            "“", "’", "”", "‘", "÷", "ÿ", "€", "Ÿ", "⁄", "‹", "ﬂ", "‚", "‡", _
            "·", "„", "ö", "‰", "Á", "Â", "Ê", "Ë", "õ", "È", "Ï", "Í", "Î", _
            "Ì", "Ó", "Ú", "Ô", "Ò", "Û", "ú", "Ù", "˜", "ı", "ˆ", "¯", "˘", _
            "¸", "˙", "˚", "ˇ")
            
            
End Sub



Public Function IsDoubleCharSet(cdtbl As Integer) As Boolean
    Select Case cdtbl
        Case VIQR_TABLE_ENUM, VNI_WINDOWS_TABLE_ENUM, VNCP_1258_TABLE_ENUM, UNICODE_COMPOSED_TABLE_ENUM, UTF8_TABLE_ENUM, BKHCM2_TABLE_ENUM
            IsDoubleCharSet = True
        Case Else
            IsDoubleCharSet = False
    End Select
End Function



Public Sub SetCodeTable(codetbl As Integer)
    Dim I As Integer
    If Not CodeTableMaked Then MakeCodeTable
    CodeTable = codetbl
    Select Case CodeTable
        Case 1: frmMain.lstCode.ListIndex = 0
        Case 3: frmMain.lstCode.ListIndex = 1
        Case 4: frmMain.lstCode.ListIndex = 2
        Case 5: frmMain.lstCode.ListIndex = 3
        Case 6: frmMain.lstCode.ListIndex = 4
        Case 8: frmMain.lstCode.ListIndex = 5
        Case 9: frmMain.lstCode.ListIndex = 6
        Case 10: frmMain.lstCode.ListIndex = 7
        Case 11: frmMain.lstCode.ListIndex = 8
        Case 12: frmMain.lstCode.ListIndex = 9
        Case 13: frmMain.lstCode.ListIndex = 10
        Case 15: frmMain.lstCode.ListIndex = 11
    End Select
    frmMenu.mnucode(CodeTable).Checked = True
    For I = 1 To frmMenu.mnucode.COunt
        If I <> CodeTable Then frmMenu.mnucode(I).Checked = False
    Next I
End Sub


Public Function CodeTableConvert(stringIn As String, Optional codeIn As Integer = 1, Optional codeOut As Integer = 1) As String
    If stringIn = "" Then Exit Function
    Dim I As Long, J As Long, S As String, sTemp As String
    sTemp = stringIn
    If codeIn = codeOut Then GoTo ENDSOON
    If Not IsDoubleCharSet(codeIn) Then
        For I = 1 To Len(stringIn)
            S = Mid$(stringIn, I, 1)
            For J = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
                If codeIn = BKHCM1_TABLE_ENUM Then
                    If codeOut = BKHCM2_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                ElseIf codeIn = BKHCM2_TABLE_ENUM Then
                    If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = TCVN3_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_COMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UTF8_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIETWARE_F_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIQR_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VISCII_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNCP_1258_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNI_WINDOWS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VPS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    End If
                
                End If
            Next J
        Next I
    Else
        
        For I = 1 To Len(stringIn)
            'Lay 3 ky tu de xu ly
            S = Mid$(stringIn, I, 3)
            For J = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
                If codeIn = BKHCM1_TABLE_ENUM Then
                    If codeOut = BKHCM2_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                ElseIf codeIn = BKHCM2_TABLE_ENUM Then
                    If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = TCVN3_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_COMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UTF8_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIETWARE_F_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIQR_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VISCII_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNCP_1258_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNI_WINDOWS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VPS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    End If
                
                End If
            Next J
        Next I
        
        ' Lay 2 ky tu de xu ly
        For I = 1 To Len(stringIn)
            S = Mid$(stringIn, I, 2)
            For J = UBound(UNICODE_PRECOMPOSED_TABLE) To LBound(UNICODE_PRECOMPOSED_TABLE) Step -1
                If codeIn = BKHCM1_TABLE_ENUM Then
                    If codeOut = BKHCM2_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM1_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                ElseIf codeIn = BKHCM2_TABLE_ENUM Then
                    If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = BKHCM2_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = TCVN3_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = TCVN3_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_COMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_COMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UNICODE_PRECOMPOSED_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = UTF8_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = UTF8_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIETWARE_F_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIETWARE_F_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VIQR_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VIQR_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VISCII_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VISCII_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNCP_1258_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNCP_1258_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VNI_WINDOWS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VPS_TABLE_ENUM Then
                        If S = VNI_WINDOWS_TABLE(J) Then sTemp = Replace$(sTemp, S, VPS_TABLE(J))
                    End If
                
                ElseIf codeIn = VPS_TABLE_ENUM Then
                   If codeOut = BKHCM1_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM1_TABLE(J))
                    ElseIf codeOut = BKHCM2_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, BKHCM2_TABLE(J))
                    ElseIf codeOut = TCVN3_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, TCVN3_TABLE(J))
                    ElseIf codeOut = UNICODE_COMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_COMPOSED_TABLE(J))
                    ElseIf codeOut = UNICODE_PRECOMPOSED_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UNICODE_PRECOMPOSED_TABLE(J))
                    ElseIf codeOut = UTF8_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, UTF8_TABLE(J))
                    ElseIf codeOut = VIETWARE_F_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIETWARE_F_TABLE(J))
                    ElseIf codeOut = VIQR_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VIQR_TABLE(J))
                    ElseIf codeOut = VISCII_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VISCII_TABLE(J))
                    ElseIf codeOut = VNCP_1258_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNCP_1258_TABLE(J))
                    ElseIf codeOut = VNI_WINDOWS_TABLE_ENUM Then
                        If S = VPS_TABLE(J) Then sTemp = Replace$(sTemp, S, VNI_WINDOWS_TABLE(J))
                    End If
                
                End If
            Next J
        Next I
    
    End If
ENDSOON:
    CodeTableConvert = sTemp
End Function
