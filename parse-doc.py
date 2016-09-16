import docx
import parse_sections

def format_title(text, i, includeLine = False):
    return format_internal(text, 1, i, includeLine)

def format_header(text, i, includeLine = False):
    return format_internal(text, 2, i, includeLine)

def format_subheader(text, i, includeLine = False):
    return format_internal(text, 3, i, includeLine)

def format_internal(text, level, i, includeLine):
    text = text.strip()
    if(text):
        prefix = "#" * level
        if(includeLine):
            return "%s (%d) %s" % (prefix, i, text)
        else:
            return "%s %s" % (prefix, text)
    else:
        return ""

    
path = "/Users/rhaffey/Dropbox/Projects/EMS/pdf-work/guidelines.docx"

doc = docx.Document(path)

start = 0
p_count = len(doc.paragraphs) #100
h_log = open('headings_X.txt', 'w')

def log(text):
    global h_log
    h_log.write(text)
    h_log.write("\n")

h = set()
h2 = set()

for i in range(start, start + p_count):
    p = doc.paragraphs[i]
    style = p.style.name
    h.add(style)

    s = ""
    if("Title" in style):
        s = format_title(p.text, i)
    elif("Heading1" in style or "Heading 1" in style):
        s = format_header(p.text, i)
    elif("Heading2" in style or "Heading 2" in style):
        s = format_subheader(p.text, i)

    if(s):
        print(s)
        log(s)

    runheading = ""
    for r in p.runs:
        style = r.style.name
        if("Heading" in style):
            runheading = runheading + r.text

    if(runheading.strip()):
        s = format_subheader(runheading, i)
        print(s)
        log(s)

    section = parse_sections.get_section_header(p)
    if(section):
        print("  * " + section)

h_log.close()        


    
