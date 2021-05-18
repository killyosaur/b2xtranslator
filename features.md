# Office Binary (doc, xls, ppt) Translator to Open XML by DIaLOGIKa

## Menu

* [About](./README.md)
* [User Documentation](./documentation.md)
* [Developer's Corner](./architecture.md)
* [Supplementary Downloads](./download.md)

### Table of Contents

* [doc to docx Translator](#doc-to-docx-translator)
* [xls to xlsx Translator](#xls-to-xlsx-translator)
* [ppt to pptx Translator](#ppt-to-pptx-translator)

### doc to docx Translator

The doc to docx translator is already quite mature. Nevertheless, some of the rather infrequently used features are not yet or not yet fully implemented. The following tasks might be tackled in future phases of this project:

* improving charts, SmartArts and WordArts translation
* improving revision mark translation

Please report any problems or issues you encounter in the SourceForge ["doc2x Maintenance"](https://sourceforge.net/tracker2/?atid=1126159&group_id=216787&func=browse) tracker.

### xls to xlsx Translator

The current goal of the xls to xlsx translator is to retain as many of the Excel-typical features as possible like formulas, value and date formatting, string and cell formatting.
The conversion result of straightforward Excel sheets should be quite good. More complex features are currently not supported, such as:

* Charts
* Pivot Tables
* Pictures
* Shapes
* OLE Objects
* Comments
* Macros
* Document Properties

You can submit bugs or feature requests on Sourceforge in the ["xls2x Bugs"](https://sourceforge.net/tracker2/?atid=1038368&group_id=216787&func=browse) tracker.

### ppt to pptx Translator

The ppt to pptx translator is in an development phase where we retain elements that are important for the semantics of a presentation. Such elements are slides, text on slides, numbered and bulleted lists, pictures, some shapes as well as basic animations.
Some features that are not yet implemented:

* Charts
* Tables
* Comments
* Macros
* Some Animationen
* Document Properties
* OLE Objects

Anyhow feel free to submit bugs in the ["ppt2x Bugs"](https://sourceforge.net/tracker2/?atid=1126169&group_id=216787&func=browse) tracker on sourceforge

### Bugs

**Note from Killyosaur** this is a copy/paste of the sourceforge issues page, I may investigate fixing issues and updating this process, so if you do find issues, put them in the main issues for this repo. Thanks!