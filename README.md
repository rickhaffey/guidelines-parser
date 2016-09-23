# guidelines-parser

Scripts to parse the [National EMS Clinical Guidelines](https://nasemso.org/Projects/ModelEMSClinicalGuidelines/documents/National-Model-EMS-Clinical-Guidelines-23Oct2014.pdf) into a more structured format.

## Goals

* The goal of these scripts is to parse the documents (from either .pdf or .docx formats) into some structured format (.json) that can then be more easily accessed from a programmatic standpoint.
* Broader goals (underlying desire for a more 'strucured' format for the content), include things like:
	* Performing text analysis of guidelines against published research, etc. 
    * Supporting applications / systems for extending the doc, collaborating on changes, etc.
    * Providing multiple representations of the document (i.e. relationships between guidelines, flowcharts generated from doc content, different views for different consumers, etc.)
    * Linking to external references and related materials (i.e. link to nemsis, ems compass documentation, external images and content, etc.)

## Notes

* Rather than parsing the PDF, I've opted to instead use the Word (.docx) version of the document, in that it provides more context as to the document structure.
* Performance is of lower concern here -- the number of runs should be minimal
