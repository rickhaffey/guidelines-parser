import docx
import parse_sections
import json

def build_category(text):
    return ("category", text.strip())

def build_guideline(text, i, category):
    return ("guideline", {
        "title": text.strip(),
        "category": category,
        "sections": [],
        "idx": i
    })

def parse_paragraph(p, i, category = None):
    style = p.style.name

    if("Heading1" in style or "Heading 1" in style):
        return build_category(p.text)
    elif("Heading2" in style or "Heading 2" in style):
        return build_guideline(p.text, i, category)

    runheading = ""
    for r in p.runs:
        style = r.style.name
        if("Heading" in style):
            runheading = runheading + r.text

    if(runheading.strip()):
        return build_guideline(runheading, category)

    return None


def get_ppr(p):
  """Gets the product properties (serialized as <pPr> in the .docx xml) for the paragraph"""
  return p._element.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")

def get_ilvl(ppr):
  """Gets the indent level (serialized as <pPr><numPr><ilvl val="_"> in the .docx xml) from the product properties"""
  if(ppr is None or ppr.numPr is None or ppr.numPr.ilvl is None or ppr.numPr.ilvl.val is None):
    return 0
  return int(ppr.numPr.ilvl.val) + 1


def main():
#    path = "/Users/rhaffey/Dropbox/Projects/EMS/pdf-work/guidelines.docx"
    path = "/Users/rhaffey/Dropbox/Projects/EMS/pdf-work/guidelines.partial.docx"
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

        parsed = parse_paragraph(p, i, category)
        if(parsed):
            if(parsed[0] == "category"):
                category = parsed[1]
                print('#', end='', flush=True)
            elif(parsed[0] == "guideline"):
                guideline = parsed[1]
                out_doc["guidelines"].append(guideline)
                print('+', end='', flush=True)

        section = parse_sections.get_section_header(p)
        if(section):
            running_section = { "heading": section, "text": [], "idx": i }
            guideline["sections"].append(running_section)
        elif(parsed == None and running_section != None):
            outText = p.text.strip()
            if(outText != ""):
                print(".", end="", flush=True)
                ppr = get_ppr(p)
                ilvl = get_ilvl(ppr)
                running_section["text"].append({ "text": p.text, "lvl": ilvl, "idx": i })
            else:
                print("x", end="", flush=True)

    with open('output/out.indented.json', 'w') as outfile:
        outfile.write(json.dumps(out_doc, indent=2))


if __name__ == '__main__':
    main()
