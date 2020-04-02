
from cdr import CDR

docPath = 'D:\\AILab\\Data\\CDRs\\test.docx'
TEST = 'C:\\Users\\Administrator\\Desktop\\test.cdr'


cdr = CDR(TEST)
cdr.togglePage(1)

layer = cdr.doc.ActiveLayer
shape = cdr.doc.ActiveShape

cdr.togglePage(1)
story = shape.Text.Story
paragraphs = story.paragraphs

modifyParagraph = []

for paragraph in paragraphs:
    thetext = paragraph.WideText
    if '\t' in thetext:
        # print('thetext',thetext)
        splits = thetext.split('\t')
        prefix = splits[0]
        theord = ord(prefix[0])
        if theord > 61000:
            remain = ''
            if len(splits) > 0:
                remain = splits[1]


            paragraph.Text = remain.strip() + "\r" 
            paragraph.ApplyBulletEffect(prefix, None, paragraph.Size, -1)
          

            # modifyParagraph.append({
            #     'remain':remain,
            #     'prefix':prefix,
            #     'paragraph':paragraph
            # })
            pass
    pass

# for p in modifyParagraph:
#     paragraph = p['paragraph']
#     prefix = p['prefix']
#     remain = p['remain']
#     paragraph.ApplyBulletEffect(prefix, None, paragraph.Size, -1)
#     # paragraph.Text = remain.replace('\r','',1)



'''thetables = doc.tables
for table in thetables:
    therows = table.rows

    for row in therows:
        thecells = row.cells
        for cell in thecells:
            theparagraphs = cell.paragraphs
            for p in theparagraphs:
                thetext = p.text
                gettag = thetext
            pass'''