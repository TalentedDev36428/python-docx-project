###########
Python DocX
###########

*Python DocX* is now part of Python OpenXML**. There's all kinds of new stuff, including Python 3 support, sister libraries for doing Excel files, and more. 

Introduction
============

The docx module creates, reads and writes Microsoft Office Word 2007 docx
files.

These are referred to as 'WordML', 'Office Open XML' and 'Open XML' by
Microsoft.

These documents can be opened in Microsoft Office 2007 / 2010, Microsoft Mac
Office 2008, Google Docs, OpenOffice.org 3, and Apple iWork 08.

They also `validate as well formed XML <http://validator.w3.org/check>`_.

The module was created when I was looking for a Python support for MS Word
.docx files, but could only find various hacks involving COM automation,
calling .Net or Java, or automating OpenOffice or MS Office.

The docx module has the following features:

Making documents
----------------

Features for making documents include:

- Paragraphs
- Bullets
- Numbered lists
- Document properties (author, company, etc)
- Multiple levels of headings
- Tables
- Section and page breaks
- Images


Editing documents
-----------------

Thanks to the awesomeness of the lxml module, we can:

- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - useful for
  retrieving data from user-completed templates.


Getting started
===============

Making and Modifying Documents
------------------------------

- clone 
- Use **pip** or **easy_install** to fetch the **lxml** and **PIL** modules.
- Then run::

    example-makedocument.py


Congratulations, you just made and then modified a Word document!


Extracting Text from a Document
-------------------------------

If you just want to extract the text from a Word file, run::

    example-extracttext.py 'Some word file.docx' 'new file.txt'

