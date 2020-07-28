from PIL import Image,ImageDraw,ImageFont
from fontTools.ttLib import TTFont
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable


def getTextWidth(text,pointSize):
    font = TTFont('c:/simsun.ttf')
    cmap = font['cmap']
    t = cmap.getcmap(3,1).cmap
    s = font.getGlyphSet()
    units_per_em = font['head'].unitsPerEm

    total = 0
    for c in text:
        if ord(c) in t and t[ord(c)] in s:
            total += s[t[ord(c)]].width
        else:
            total += s['.notdef'].width
    total = total*float(pointSize)/units_per_em;
    return total


def test():
  width = getTextWidth('的',10)
  print(width)


if __name__ == '__main__':
  test()
