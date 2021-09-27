# Office Binary (doc, xls, ppt) Translator to Open XML by DIaLOGIKa

Transferred this from the SourceForce.net site, because reasons. [B2XTranslator](http://b2xtranslator.sourceforge.net/index.html)

## Menu

* [User Documentation](./documentation.md)
* [Developer's Corner](./architecture.md)
* [Known Issues](./features.md)
* [Supplementary Downloads](./download.md)

### Table of Contents

* [Overview](#overview)
* [Main Contributors](#main-contributors)
* [Licensing model](#licensing-model)

### Overview

The main goal of the Office Binary (doc, xls, ppt) Translator to Open XML project is to create software tools, plus guidance, showing how a document written using the Binary Formats (doc, xls, ppt) can be translated into the Office Open XML format (aka OpenXML). As a result customers can use these tools to migrate from the binary formats to OpenXML; thus, enabling them to more easily access their existing content in the new world of XML. The Translator will be available under the open source Berkeley Software Distribution (BSD) [license](#licensing-model), which allows that anyone can use the mapping and code, submit bugs and feedback, or contribute to the project.

On February 15th 2008, Microsoft has made it even easier to get [access to the **binary formats documentation**](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22). This documentation has been completely revamped by June 30, 2008 and the latest version can be [downloaded from here](https://interoperability.blob.core.windows.net/files/MS-DOC/%5bMS-DOC%5d.pdf). All these specifications are available under the [Open Specification Promise](http://www.microsoft.com/interop/osp).

The **Office Open XML file formats** have been approved and [published](http://www.iso.org/iso/pressrelease.htm?refid=Ref1181) as ISO/IEC 29500. The pre-ISO version or ECMA-376 1st Edition, which is implemented in Office 2007 SP2, and ECMA-376 2nd edition, which is technically aligned with ISO/IEC 29500, are available [free of charge from Ecma-International](http://www.ecma-international.org/publications/standards/Ecma-376.htm). ISO/IEC 29500 can be purchased from [ISO/IEC](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463).

Microsoft also published a set of document-format implementation notes for ECMA-376 1st Edition. The goal of publishing these notes is to help other implementers improve interoperability with Office, by transparently documenting the details of Microsoft's OpenXML implementation. To get to the ECMA-376 implementer notes, go to the [DII home page](http://www.documentinteropinitiative.org/) and click on Reference and then select ECMA-376 1st Edition from the dropdown list.

While Microsoft provides with Office 2007 and the File Format Compatibility pack for earlier Office versions a migration path from binary Office formats to OpenXML, the Office Binary (doc, xls, ppt) Translator to Open XML project is still necessary due to the following reasons

* Enables the back-office / batch scenario due to its a command-line-based architecture
* Provides a cross-platform story via .Net/Mono, i.e. it the translators run, for example, on SUSE Linux
* Proves the usability and completeness of the file format specifications
* Allows that anyone uses the mapping, code snippets, etc. due to the open source development approach based on the liberate BSD license

We have chosen to use an Open Source development model that allows developers from all around the world to participate and contribute to the project.

### Main Contributors

* **[DIaLOGIKa](http://www.dialogika.de/) (Analysis and Development)**

  DIaLOGIKa - a German systems and software house founded in 1982 - conducts projects on behalf of industry, finance, and governmental and supranational clients such as the institutions of the European Union (EU).

  From the beginning DIaLOGIKa has focused – among others – on technically demanding projects in the field of multilingual text and data processing such as document format conversion. DIaLOGIKa has also contributed to the [OpenXML/ODF Translator](https://sourceforge.net/projects/odf-converter/) project.

* **[Microsoft](http://www.microsoft.com/interop) (Architectural Guidance, Technical Support & Project Management)**

### Licensing Model

This project is developed and released under a very liberal BSD-like license:

``` LICENSE
* Copyright (c) 2008-2009, DIaLOGIKa
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of DIaLOGIKa nor the
*       names of its contributors may be used to endorse or promote products
*       derived from this software without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```
