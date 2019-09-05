import sys
import chardet
from fontTools import ttLib

FONTNAME = "华康娃娃体W5.ttf"

FONT_SPECIFIER_PSNAME_ID = 6
FONT_SPECIFIER_NAME_ID = 4
FONT_SPECIFIER_FAMILY_ID = 1
CHINESELANGUAGE = 33


def convertUnicode(namestr):
    res = chardet.detect(namestr)
    if (res['encoding'] == 'ascii'):
        return namestr.decode('ascii')
    else:
        codec = res['encoding']
        if codec is None:
            codec = 'GB2312'
        return namestr.decode(codec)


def shortName(font):
    """Get the short name from the font's names table"""
    postscriptName = ""
    familyName = ""
    chineseName = ""
    for record in font['name'].names:
        if record.langID == 0:
            if record.nameID == FONT_SPECIFIER_PSNAME_ID and not postscriptName:
                postscriptName = convertUnicode(record.string)
            elif record.nameID == FONT_SPECIFIER_FAMILY_ID and not familyName:
                familyName = convertUnicode(record.string)
        elif record.langID == CHINESELANGUAGE:
            if record.nameID == FONT_SPECIFIER_NAME_ID and not chineseName:
                chineseName = convertUnicode(record.string)
        if postscriptName and familyName and chineseName:
            break
    return postscriptName, familyName, chineseName


tt = ttLib.TTFont(FONTNAME)
theName = shortName(tt)
humanName = theName[1]
if (theName[2] != ''):
    humanName = theName[2]
thefilename = theName[0] + '(' + humanName + ')'
print(thefilename)
