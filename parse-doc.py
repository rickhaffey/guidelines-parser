import docx
import json
import re
import sys

default_root = "/Users/rhaffey/Dropbox/Projects/EMS/pdf-work"

json_category = "category"
json_guidelines = "guidelines"
json_guideline = "guideline"
json_title = "title"
json_sections = "sections"
json_p_index = "p_index"
json_heading = "heading"
json_text = "text"
json_indent = "indent"

sectionRegexes = {
  r"""secondary.*assessment.*treatment.*interventions""": "Secondary Assessment, Treatment, and Interventions",
  r"""assessment.*treatment.*interventions""": "Assessment, Treatment, and Interventions",
  r"""treatment.*interventions""": "Treatment and Interventions",
  r"""assessment""": "Assessment",
  r"""definitions""": "Definitions",
  r"""inclusion.*exclusion.*criteria""": "Inclusion / Exclusion Criteria",
  r"""exclusion.*criteria""": "Exclusion Criteria",
  r"""inclusion.*criteria""": "Inclusion Criteria",
  r"""key.*considerations""": "Key Considerations",
  r"""key.*documentation.*elements""": "Key Documentation Elements",
  r"""notes.*educational.*pearls""": "Notes / Educational Pearls",
  r"""patient.*care.*goals""": "Patient Care Goals",
  r"""patient.*management""": "Patient Management",
  r"""patient.*presentation""": "Patient Presentation",
  r"""patient.*safety.*considerations""": "Patient Safety Considerations",
  r"""performance.*measures""": "Performance Measures",
  r"""pertinent.*assessment.*findings""": "Pertinent Assessment Findings",
  r"""quality.*improvement""": "Quality Improvement",
  r"""references""": "References",
  r"""special.*transport.*considerations""": "Special Transport Considerations",
  r"""scene.*management""": "Scene Management"
}

def build_category(text):
    return (json_category, text.strip())


def build_guideline(text, i, category):
    return (json_guideline, {
        json_title: text.strip(),
        json_category: category,
        json_sections: [],
        json_p_index: i
    })


def build_section(text, i):
    return {
        json_heading: text,
        json_text: [],
        json_p_index: i
    }


def build_section_text(text, i, indent):
    return {
        json_text: text,
        json_indent: indent,
        json_p_index: i
    }


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
        return build_guideline(runheading, i, category)

    return None


def parse_section_header(p):
    rtext = ""
    isSection = False
    for r in p.runs:
        rtext = rtext + r.text.strip()
        f = r.font
        isSection = isSection or (f.bold and f.underline)

    if(isSection):
        # check to see whether it matches one of our regexes
        for regex in sectionRegexes.keys():
            if(re.match(regex, rtext, re.I)):
                return sectionRegexes[regex]

    return None


def get_ppr(p):
  """Gets the product properties (serialized as <pPr> in the .docx xml) for the paragraph"""
  return p._element.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")


def get_ilvl(p):
  """Gets the indent level (serialized as <pPr><numPr><ilvl val="_"> in the .docx xml) from the product properties"""
  ppr = get_ppr(p)

  if(ppr is None or ppr.numPr is None or ppr.numPr.ilvl is None or ppr.numPr.ilvl.val is None):
    return 0
  return int(ppr.numPr.ilvl.val) + 1


def show_progress(marker):
    print(marker, end='', flush=True)


def get_infile_path():
    if(len(sys.argv) >= 2):
        return sys.argv[1]
    else:
        return default_root + "/guidelines.docx"
        #return default_root + "/guidelines.partial.docx"


def get_outfile_path():    
    if(len(sys.argv) >= 3):
        return sys.argv[2]
    else:
        return default_root + "/output/out.json"

    
def main():
    # read in the input word doc
    path = get_infile_path()
    source_doc = docx.Document(path)

    # build the starting point of an output doc
    out_doc = {
        json_guidelines: []
    }

    # build some preliminary 'running' objects
    running_category = None
    running_guideline = None
    running_section = None

    # iterate over the paragraphs in the doc
    start = 0
    end = start + len(source_doc.paragraphs)
    for i in range(start, end):
        p = source_doc.paragraphs[i]

        # first, handle new category and guideline paragraphs
        parsed = parse_paragraph(p, i, running_category)
        if(parsed):
            if(parsed[0] == "category"):
                running_category = parsed[1]
                show_progress("#")
            elif(parsed[0] == "guideline"):
                running_guideline = parsed[1]
                out_doc[json_guidelines].append(running_guideline)
                show_progress("+")

        # next, parse out any section details (headers, and section text)
        section = parse_section_header(p)
        if(section):
            running_section = build_section(section, i)
            running_guideline[json_sections].append(running_section)
        elif(parsed == None and running_section != None):
            outText = p.text.strip()
            if(outText != ""):
                section_text = build_section_text(p.text, i, get_ilvl(p))
                running_section["text"].append(section_text)
                show_progress(".")
            else:
                show_progress("x")

    # write the final document in .json format to the outfile
    with open(get_outfile_path(), 'w') as outfile:
        outfile.write(json.dumps(out_doc, indent=2))


if __name__ == '__main__':
    main()
