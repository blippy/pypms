# create html file

import glob
import os

import common


template = '''
<html>
<head> <title>Monthly reports</title> </head>
<body>
<p>A list of text files available</p>

CONTENTS

</body>
</html>
'''



def create_html():
    global template
    root = common.outroot #"m:\\Finance\\pypms\\out"
    nodes = os.listdir(root)
    dirs = [node for node in nodes if os.path.isdir(root+ "\\" + node)]
    dirs.sort()
    dirs.reverse()
    contents = ''
    for d in dirs:
        contents += "<p>{0}: ".format(d) 
        txtfiles = glob.glob(root + "\\" + d + "\\reports\\*.txt")
        for txtfile in txtfiles:
            name_without_suffix = os.path.splitext(os.path.basename(txtfile))[0]
            href = '<a href="file://{0}">{1}</a> '.format(txtfile, name_without_suffix)
            contents += href
        contents += "</p>\n\n"
    html = template.replace('CONTENTS', contents)
    #file(root + "\\texts.htm", 'w').write(html)
    common.spit(root + "\\texts.htm", html)

if __name__ == "__main__":
    create_html()
    print 'Finished'
