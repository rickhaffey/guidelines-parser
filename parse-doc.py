import docx
import parse_sections
import json

def build_category(text):
    return ("category", text.strip())

def build_guideline(text, category):
    return ("guideline", {
        "title": text.strip(),
        "category": category,
        "sections": []
    })

# def format_header(text, i, includeLine = False):
#     return format_internal(text, 1, i, includeLine)

# def format_subheader(text, i, includeLine = False):
#     return format_internal(text, 2, i, includeLine)

# def format_internal(text, level, i, includeLine):
#     text = text.strip()
#     if(text):
#         prefix = "#" * level
#         if(includeLine):
#             return "%s (%d) %s" % (prefix, i, text)
#         else:
#             return "%s %s" % (prefix, text)
#     else:
#         return ""

def parse_paragraph(p, category = None):
    style = p.style.name

    if("Heading1" in style or "Heading 1" in style):
        return build_category(p.text)
    elif("Heading2" in style or "Heading 2" in style):
        return build_guideline(p.text, category)

    runheading = ""
    for r in p.runs:
        style = r.style.name
        if("Heading" in style):
            runheading = runheading + r.text

    if(runheading.strip()):
        return build_guideline(runheading, category)

    return None


def main():
    path = "/Users/rhaffey/Dropbox/Projects/EMS/pdf-work/guidelines.docx"
    source_doc = docx.Document(path)
    
    start = 0
    p_count = len(source_doc.paragraphs)

    out_doc = {
        "guidelines": []
    }

    category = None
    guideline = None
    running_section = None
    for i in range(start, start + p_count):
        p = source_doc.paragraphs[i]
    
        parsed = parse_paragraph(p, category)
        if(parsed):
            if(parsed[0] == "category"):
                category = parsed[1]
            elif(parsed[0] == "guideline"):
                guideline = parsed[1]
                out_doc["guidelines"].append(guideline)

            print('.', end='', flush=True)
    
        section = parse_sections.get_section_header(p)
        if(section):
            running_section = { "heading": section, "text": [] }
            guideline["sections"].append(running_section)
        elif(parsed == None and running_section != None):
            print("+", end="", flush=True)
            running_section["text"].append(p.text)

    with open('output/out.json', 'w') as outfile:
        outfile.write(json.dumps(out_doc))
    

if __name__ == '__main__':
    main()
